using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Tesseract;
using SoborSmartReport.Ocr;

namespace SoborSmartReport
{
    public class QualityPdfReader
    {
        private readonly OcrReader _ocr;

        public QualityPdfReader(
            OcrReader ocr)
        {
            _ocr = ocr ??
                throw new ArgumentNullException(nameof(ocr));
        }

        /// <summary>
        /// PDF를 렌더링한 Bitmap을 인식한다.
        /// </summary>
        public QualityOcrResult Read(
            Bitmap pageImage,
            string debugOutputDir = null)
        {
            if (pageImage == null)
                throw new ArgumentNullException(nameof(pageImage));

            QualityOcrResult result =
                new QualityOcrResult();

            foreach (QualityOcrRegion region
                in CreateRegions())
            {
                QualityOcrItem item =
                    ReadRegion(
                        pageImage,
                        region,
                        debugOutputDir);

                result.Items.Add(
                    region.Name,
                    item);
            }

            return result;
        }

        private QualityOcrItem ReadRegion(
            Bitmap pageImage,
            QualityOcrRegion region,
            string debugOutputDir)
        {
            Rectangle pixelRect =
                OcrImageProcessor.GetRatioRectangle(
                    pageImage,
                    region.X,
                    region.Y,
                    region.Width,
                    region.Height);

            try
            {
                using (Bitmap crop =
                    pageImage.Clone(
                        pixelRect,
                        pageImage.PixelFormat))
                {

#if OPENCV

                    using (Bitmap processed =
                        OcrImageProcessor.Preprocess(
                            crop,
                            region.Options.Scale,
                            region.Options.Grayscale,
                            region.Options.Threshold))
                    {
#if DEBUG
                        SaveDebugImage(
                            processed,
                            region.Name,
                            debugOutputDir);
#endif

                        OcrReadResult ocrResult =
                            _ocr.Read(
                                processed,
                                region.PageSegMode);

                        return CreateResult(
                            region,
                            ocrResult,
                            pixelRect);
                    }

#else

                    OcrReadResult ocrResult =
                        _ocr.Read(
                            crop,
                            region.PageSegMode);

                    return CreateResult(
                        region,
                        ocrResult,
                        pixelRect);

#endif
                }
            }
            catch (Exception ex)
            {
                return new QualityOcrItem
                {
                    Name = region.Name,
                    Success = false,
                    PixelRectangle = pixelRect,
                    Error = ex.Message
                };
            }
        }

        private QualityOcrItem CreateResult(
            QualityOcrRegion region,
            OcrReadResult ocrResult,
            Rectangle pixelRect)
        {
            string normalized =
                Normalize(
                    region,
                    ocrResult.Text);

            double? numericValue = null;

            if (region.IsNumeric)
            {
                if (double.TryParse(
                    normalized,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out double value))
                {
                    numericValue = value;
                }
            }

            bool success;

            if (region.IsNumeric)
            {
                success =
                    numericValue.HasValue &&
                    ocrResult.Confidence >=
                        region.MinConfidence;
            }
            else
            {
                success =
                    !string.IsNullOrWhiteSpace(normalized) &&
                    ocrResult.Confidence >=
                        region.MinConfidence;
            }

            bool formatValid = true;

            if (region.Name == "measurement_period")
            {
                normalized = NormalizeMeasurementPeriod(
                    ocrResult.Text,
                    out formatValid);
            }
            else
            {
                normalized = Normalize(
                    region,
                    ocrResult.Text);
            }

            return new QualityOcrItem
            {
                Name = region.Name,

                Success = success,

                RawText = ocrResult.Text,

                Text = normalized,

                FormatValid = formatValid,

                NumericValue = numericValue,

                Confidence =
                    ocrResult.Confidence,

                PixelRectangle =
                    pixelRect
            };
        }

        private string Normalize(
            QualityOcrRegion region,
            string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            text = text.Trim();

            // 줄바꿈 제거
            text = text
                .Replace("\r", " ")
                .Replace("\n", " ");

            // 연속 공백
            text = Regex.Replace(
                text,
                @"\s+",
                " ");

            if (region.IsNumeric)
            {
                // 숫자, 점, 마이너스만 유지
                Match match = Regex.Match(
                    text,
                    @"-?\d+(?:\.\d+)?");

                if (match.Success)
                    return match.Value;

                return "";
            }

            return text.Trim();
        }

#if DEBUG

        private void SaveDebugImage(
            Bitmap bitmap,
            string name,
            string outputDir)
        {
            if (string.IsNullOrWhiteSpace(
                outputDir))
            {
                return;
            }

            Directory.CreateDirectory(
                outputDir);

            string file =
                Path.Combine(
                    outputDir,
                    $"{name}.png");

            bitmap.Save(
                file,
                System.Drawing.Imaging.ImageFormat.Png);
        }

#endif

        // ==========================================================
        // sample.pdf ROI
        // ==========================================================

        private List<QualityOcrRegion>
            CreateRegions()
        {
            OcrProcessOptions normal =
                new OcrProcessOptions
                {
                    Scale = 3.0,
                    Grayscale = true,
                    Threshold = "otsu"
                };

            OcrProcessOptions text =
                new OcrProcessOptions
                {
                    Scale = 4.0,
                    Grayscale = true,
                    Threshold = "otsu"
                };

            return new List<QualityOcrRegion>
            {
                // -----------------------------------------
                // 측정 기간
                // -----------------------------------------

                new QualityOcrRegion(
                    "measurement_period",
                    0.4,
                    0.155,
                    0.45,
                    0.02,
                    text,
                    false,
                    PageSegMode.SingleLine),

                // -----------------------------------------
                // 기록 간격
                // -----------------------------------------

                new QualityOcrRegion(
                    "record_interval",
                    0.14,
                    0.181,
                    0.2,
                    0.025,
                    text,
                    false,
                    PageSegMode.SingleLine),

                // -----------------------------------------
                // 최대전력 / 역률
                // -----------------------------------------

                new QualityOcrRegion(
                    "max_power",
                    0.16, 0.36,
                    0.07, 0.07,
                    normal),

                new QualityOcrRegion(
                    "power_factor",
                    0.25, 0.36,
                    0.065, 0.07,
                    normal),

                // -----------------------------------------
                // 전압
                // -----------------------------------------

                new QualityOcrRegion(
                    "voltage_r",
                    0.363, 0.348,
                    0.07, 0.02,
                    normal),

                new QualityOcrRegion(
                    "voltage_s",
                    0.363, 0.388,
                    0.07, 0.02,
                    normal),

                new QualityOcrRegion(
                    "voltage_t",
                    0.363, 0.428,
                    0.07, 0.02,
                    normal),

                // -----------------------------------------
                // 전압 THD
                // -----------------------------------------

                new QualityOcrRegion(
                    "voltage_thd_r",
                    0.455, 0.340,
                    0.060, 0.030,
                    normal),

                new QualityOcrRegion(
                    "voltage_thd_s",
                    0.455, 0.380,
                    0.060, 0.030,
                    normal),

                new QualityOcrRegion(
                    "voltage_thd_t",
                    0.455, 0.420,
                    0.060, 0.030,
                    normal),

                // -----------------------------------------
                // 전류
                // -----------------------------------------

                new QualityOcrRegion(
                    "current_r",
                    0.580, 0.340,
                    0.065, 0.030,
                    normal),

                new QualityOcrRegion(
                    "current_s",
                    0.580, 0.380,
                    0.065, 0.030,
                    normal),

                new QualityOcrRegion(
                    "current_t",
                    0.580, 0.420,
                    0.065, 0.030,
                    normal),

                // -----------------------------------------
                // 전류 불평형
                // -----------------------------------------

                new QualityOcrRegion(
                    "current_unbalance",
                    0.665, 0.365,
                    0.070, 0.070,
                    normal),

                // -----------------------------------------
                // 전류 THD
                // -----------------------------------------

                new QualityOcrRegion(
                    "current_thd_r",
                    0.750, 0.340,
                    0.085, 0.030,
                    normal),

                new QualityOcrRegion(
                    "current_thd_s",
                    0.750, 0.380,
                    0.085, 0.030,
                    normal),

                new QualityOcrRegion(
                    "current_thd_t",
                    0.750, 0.420,
                    0.085, 0.030,
                    normal)
            };
        }

        public QualityOcrResult Read(
            string pdfPath,
            int page = 0,
            int dpi = 400,
            string debugOutputDir = null)
        {
            if (string.IsNullOrWhiteSpace(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath));

            if (!File.Exists(pdfPath))
                throw new FileNotFoundException(
                    "PDF 파일을 찾을 수 없습니다.",
                    pdfPath);

            using (Bitmap pageImage =
                PdfImageRenderer.RenderPage(
                    pdfPath,
                    page,
                    dpi))
            {
                return Read(
                    pageImage,
                    debugOutputDir);
            }
        }

        private string NormalizeMeasurementPeriod(
            string text,
            out bool formatValid)
        {
            formatValid = false;

            if (string.IsNullOrWhiteSpace(text))
                return "";

            text = text.Trim();

            Match match = Regex.Match(
                text,
                @"^(\d{4}\.\d{2}\.\d{2})\[(\d{2}:\d{2}):\d{2}\]" +
                @"\s*~\s*" +
                @"(\d{4}\.\d{2}\.\d{2})\[(\d{2}:\d{2}):\d{2}\]$");

            // 형식이 틀리면 원문 그대로
            if (!match.Success)
                return text;

            formatValid = true;

            return
                $"{match.Groups[1].Value}    " +
                $"{match.Groups[2].Value}   ~   " +
                $"{match.Groups[3].Value}    " +
                $"{match.Groups[4].Value}";
        }
    }

    public class OcrProcessOptions
    {
        public double Scale { get; set; } = 3.0;

        public bool Grayscale { get; set; } = true;

        public string Threshold { get; set; } = "otsu";
    }

    public class QualityOcrRegion
    {
        public string Name { get; }

        public double X { get; }

        public double Y { get; }

        public double Width { get; }

        public double Height { get; }

        public OcrProcessOptions Options { get; }

        public bool IsNumeric { get; }

        public PageSegMode PageSegMode { get; }

        public float MinConfidence { get; set; } = 0.60f;

        public QualityOcrRegion(
            string name,
            double x,
            double y,
            double width,
            double height,
            OcrProcessOptions options,
            bool isNumeric = true,
            PageSegMode pageSegMode =
                PageSegMode.SingleLine)
        {
            Name = name;

            X = x;
            Y = y;
            Width = width;
            Height = height;

            Options = options;

            IsNumeric = isNumeric;

            PageSegMode = pageSegMode;
        }
    }

    public class QualityOcrItem
    {
        public string Name { get; set; }

        public bool Success { get; set; }

        /// <summary>
        /// Tesseract 원본
        /// </summary>
        public string RawText { get; set; }

        /// <summary>
        /// 정규화한 값
        /// </summary>
        public string Text { get; set; }

        public double? NumericValue { get; set; }


        public bool FormatValid { get; set; } = true;

        public float Confidence { get; set; }

        public Rectangle PixelRectangle { get; set; }

        public string Error { get; set; }
    }

    public class QualityOcrResult
    {
        public Dictionary<string, QualityOcrItem>
            Items
        { get; set; }
            =
            new Dictionary<string, QualityOcrItem>();

        public QualityOcrItem this[string name]
        {
            get
            {
                return Items.TryGetValue(
                    name,
                    out QualityOcrItem value)
                    ? value
                    : null;
            }
        }
    }
}