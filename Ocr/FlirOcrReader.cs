using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Tesseract;
using SoborSmartReport.Ocr;

#if OPENCV
using OpenCvSharp;
using OpenCvSharp.Extensions;
#endif

namespace SoborSmartReport
{
    public class FlirOcrReader : IDisposable
    {
        // private readonly TesseractEngine engine;

        private readonly OcrReader _ocr;
        private readonly bool _ownsOcr;

        public FlirOcrReader()
        {
            _ocr = new OcrReader("0123456789.");

            _ownsOcr = true;
        }

        public FlirOcrReader(string tessDataPath)
        {
            _ocr = new OcrReader(
                tessDataPath,
                "0123456789.");

            _ownsOcr = true;
        }

        public FlirOcrReader(OcrReader ocr)
        {
            _ocr = ocr ??
                throw new ArgumentNullException(nameof(ocr));

            _ownsOcr = false;
        }

        //public FlirOcrReader()
        //{
        //    string tessPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ocr\\tessdata");
        //    engine = new TesseractEngine(tessPath, "eng", EngineMode.Default);

        //    engine.SetVariable("tessedit_char_whitelist", "0123456789.");
        //}

        //public FlirOcrReader(string tessDataPath)
        //{
        //    engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default);
        //    engine.SetVariable("tessedit_char_whitelist", "0123456789.");
        //}

        public FlirResult Read(string imageFile)
        {
            using (Bitmap bmp = new Bitmap(imageFile))
            {
                FlirResult result = new FlirResult();

                using (Bitmap left = CropLeft(bmp))
                using (Bitmap right = CropRightBottom(bmp))
                using (Bitmap left2 = Preprocess(left, 8))
                using (Bitmap right2 = Preprocess(right, 6))
                {
#if DEBUG
                    string dir = Path.Combine(
                        Path.GetDirectoryName(imageFile),
                        "OcrDebug");

                    Directory.CreateDirectory(dir);


                    left.Save(Path.Combine(dir,
                        Path.GetFileNameWithoutExtension(imageFile) + "_left.png"),
                        System.Drawing.Imaging.ImageFormat.Png);

                    right.Save(Path.Combine(dir,
                        Path.GetFileNameWithoutExtension(imageFile) + "_right.png"),
                        System.Drawing.Imaging.ImageFormat.Png);

                    left2.Save(Path.Combine(dir,
                        Path.GetFileNameWithoutExtension(imageFile) + "_left_pre.png"),
                        System.Drawing.Imaging.ImageFormat.Png);

                    right2.Save(Path.Combine(dir,
                        Path.GetFileNameWithoutExtension(imageFile) + "_right_pre.png"),
                        System.Drawing.Imaging.ImageFormat.Png);
#endif

                    result.LeftRawText = Ocr(left2);
                    Debug.WriteLine($"result.LeftRawText : {result.LeftRawText}");
                    result.RightRawText = Ocr(right2);
                    Debug.WriteLine($"result.RightRawText : {result.RightRawText}");

                    result.Items = ParseItems(result.LeftRawText);
                    result.ScaleMinTemperature = ParseTemperature(result.RightRawText);
                    //result.RemoveBelowScaleMinTemperature();
                }

                return result;
            }
        }
        private string Ocr(Bitmap bmp)
        {
            OcrReadResult result =
                _ocr.Read(
                    bmp,
                    PageSegMode.SingleBlock);

            return result.Text;
        }

        private Dictionary<string, double> ParseItems(string text)
        {
            Dictionary<string, double> dic = new Dictionary<string, double>();

            //Regex regex = new Regex(@"([A-Za-z]{2}\d+).*?([\d]+\.[\d]+)");
            Regex regex = new Regex(@"\d+\.?\d*");

            double? prevValue = null;
            int index = 0;

            foreach (Match m in regex.Matches(text))
            {
                string s = NormalizeTemperature(m.Value);

                if (double.TryParse(
                    s,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out double value))
                {
                    value = CorrectTemperature(value, prevValue);
                    dic.Add(CreateItemKey(index++), value);
                    prevValue = value;
                }
            }

            return dic;
        }

        private double CorrectTemperature(double value, double? prevValue)
        {
            if (!prevValue.HasValue)
            {
                // 첫 값이 100 이상이면 앞자리 제거
                while (value >= 100)
                {
                    string tmpS = value.ToString(CultureInfo.InvariantCulture);

                    int dot = tmpS.IndexOf('.');
                    if (dot <= 1)
                        break;

                    tmpS = tmpS.Substring(1);

                    if (!double.TryParse(tmpS, NumberStyles.Any,
                        CultureInfo.InvariantCulture, out value))
                    {
                        break;
                    }
                }

                return value;
            }

            if (Math.Abs(value - prevValue.Value) <= 90)
                return value;

            string s = value.ToString(CultureInfo.InvariantCulture);

            // 첫 글자를 하나 제거해가며 확인
            while (s.Length > 3)
            {
                s = s.Substring(1);

                if (double.TryParse(s,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out double newValue))
                {
                    if (Math.Abs(newValue - prevValue.Value) <= 90)
                        return newValue;
                }
            }

            return value;
        }

        private string NormalizeTemperature(string text)
        {
            // OCR이 소수점을 놓친 경우 보정
            if (!text.Contains("."))
            {
                if (text.Length == 3)
                    text = text.Insert(2, ".");

                else if (text.Length == 4)
                    text = text.Insert(3, ".");
            }

            return text;
        }

        private string CreateItemKey(int index)
        {
            // 현재는 Item1, Item2...
            return $"Item{index + 1}";
        }

        private double ParseTemperature(string text)
        {
            Match m = Regex.Match(text, @"[\d]+\.[\d]+");

            if (!m.Success)
                return double.NaN;

            return double.Parse(m.Value, CultureInfo.InvariantCulture);
        }

        private Bitmap CropLeft(Bitmap bmp)
        {
            // 왼쪽 상단에 있는 네모 블럭을 찾아서 그 높이를 기준으로 크롭
#if OPENCV
            //try
            //{
            //    Rectangle detected = DetectTopLeftBlock(bmp);
            //    if (detected.Width > 0 && detected.Height > 0)
            //    {
            //        return bmp.Clone(detected, bmp.PixelFormat);
            //    }
            //}
            //catch
            //{
            //    // 실패하면 기존 동작으로 폴백
            //}

            try
            {
                Mat src = BitmapConverter.ToMat(bmp);
                OpenCvSharp.Rect detected = GetDynamicOverlayRoi(src);
                if (detected.Width > 0 && detected.Height > 0)
                {
                    Rectangle rect = new Rectangle(detected.Left, detected.Top,  detected.Width, detected.Height);

                    Mat overlayCrop = new Mat(src, detected);

                    Cv2.Rectangle(BitmapConverter.ToMat(bmp), detected, Scalar.Black, -1);

                    return bmp.Clone(rect, bmp.PixelFormat);
                }
            }
            catch
            {
                // 실패하면 기존 동작으로 폴백
            }

            
#endif

            Rectangle r = new Rectangle(
                (int)(bmp.Width * 0.20),
                (int)(bmp.Height * 0.01),
                (int)(bmp.Width * 0.14),
                (int)(bmp.Height * 0.38));

            return bmp.Clone(r, bmp.PixelFormat);
        }

#if OPENCV
        public static OpenCvSharp.Rect GetDynamicOverlayRoi(Mat srcImage)
        {
            // 1. 좌상단 영역(전체 이미지의 상단 50%, 좌측 50%)으로 탐색 범위 제한
            int searchWidth = srcImage.Width / 2;
            int searchHeight = srcImage.Height / 2;
            OpenCvSharp.Rect searchRoi = new OpenCvSharp.Rect(0, 0, searchWidth, searchHeight);

            Mat topLeftCrop = new Mat(srcImage, searchRoi);
            Mat gray = new Mat();
            Mat thresh = new Mat();

            // 2. 그레이스케일 변환 및 흰색 텍스트 이진화 (임계값: 220~240)
            Cv2.CvtColor(topLeftCrop, gray, ColorConversionCodes.BGR2GRAY);
            Cv2.Threshold(gray, thresh, 230, 255, ThresholdTypes.Binary);

            // 3. 윤곽선(Contour) 추출
            Cv2.FindContours(thresh, out OpenCvSharp.Point[][] contours, out HierarchyIndex[] hierarchy,
                             RetrievalModes.External, ContourApproximationModes.ApproxSimple);

            if (contours.Length == 0)
                return new OpenCvSharp.Rect(0, 0, 0, 0);

            int minX = int.MaxValue, minY = int.MaxValue;
            int maxX = int.MinValue, maxY = int.MinValue;
            bool foundText = false;

            // 4. 모든 텍스트 글자의 좌표 범위를 병합하여 전체 텍스트 영역 계산
            foreach (var contour in contours)
            {
                // 작은 노이즈 점 제외 (면적 기준)
                double area = Cv2.ContourArea(contour);
                if (area < 6) continue;

                OpenCvSharp.Rect rect = Cv2.BoundingRect(contour);

                minX = Math.Min(minX, rect.X);
                minY = Math.Min(minY, rect.Y);
                maxX = Math.Max(maxX, rect.X + rect.Width);
                maxY = Math.Max(maxY, rect.Y + rect.Height);
                foundText = true;
            }

            if (!foundText)
                return new OpenCvSharp.Rect(0, 0, 0, 0);

            // 5. 검출된 텍스트 외곽에 패딩(Margin)을 추가하여 반투명 박스까지 포함
            int padding = 8;
            int x = Math.Max(0, minX - padding);
            int y = Math.Max(0, minY - padding);
            int width = Math.Min(srcImage.Width - x, (maxX - minX) + (padding * 2));
            int height = Math.Min(srcImage.Height - y, (maxY - minY) + (padding * 2));

            return new OpenCvSharp.Rect(x, y, width, height);
        }

        // OpenCvSharp를 사용하여 왼쪽 상단의 가장 큰 사각 블럭(네모)를 검출
        private Rectangle DetectTopLeftBlock(Bitmap bmp)
        {
            using (Mat src = BitmapConverter.ToMat(bmp))
            using (Mat gray = new Mat())
            using (Mat bin = new Mat())
            {
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);

                // 왼쪽 상단만 검색
                OpenCvSharp.Rect search = new OpenCvSharp.Rect(
                    (int)(bmp.Width * 0.18),
                    0,
                    (int)(bmp.Width * 0.16),
                    (int)(bmp.Height * 0.45));

                using (Mat roi = new Mat(gray, search))
                {
                    // 흰 글자만 추출
                    Cv2.Threshold(
                        roi,
                        bin,
                        180,
                        255,
                        ThresholdTypes.Binary);

#if DEBUG
                    roi.SaveImage("roi.png");
                    bin.SaveImage("bin.png");
#endif

                    //-----------------------------------
                    // 행 방향 Projection
                    //-----------------------------------
                    int[] rowCount = new int[bin.Rows];

                    for (int y = 0; y < bin.Rows; y++)
                    {
                        int cnt = 0;

                        for (int x = 0; x < bin.Cols; x++)
                        {
                            if (bin.At<byte>(y, x) != 0)
                                cnt++;
                        }

                        rowCount[y] = cnt;
                    }

                    const int RowThreshold = 8;

                    int top = -1;
                    int bottom = -1;

                    for (int y = 0; y < rowCount.Length; y++)
                    {
                        if (rowCount[y] > RowThreshold)
                        {
                            top = y;
                            break;
                        }
                    }

                    for (int y = rowCount.Length - 1; y >= 0; y--)
                    {
                        if (rowCount[y] > RowThreshold)
                        {
                            bottom = y;
                            break;
                        }
                    }

                    if (top < 0 || bottom < top)
                        return Rectangle.Empty;

                    //-----------------------------------
                    // 열 방향 Projection
                    //-----------------------------------
                    int[] colCount = new int[bin.Cols];

                    for (int x = 0; x < bin.Cols; x++)
                    {
                        int cnt = 0;

                        for (int y = top; y <= bottom; y++)
                        {
                            if (bin.At<byte>(y, x) != 0)
                                cnt++;
                        }

                        colCount[x] = cnt;
                    }

                    const int ColThreshold = 3;

                    int left = -1;
                    int right = -1;

                    for (int x = 0; x < colCount.Length; x++)
                    {
                        if (colCount[x] > ColThreshold)
                        {
                            left = x;
                            break;
                        }
                    }

                    for (int x = colCount.Length - 1; x >= 0; x--)
                    {
                        if (colCount[x] > ColThreshold)
                        {
                            right = x;
                            break;
                        }
                    }

                    if (left < 0 || right < left)
                        return Rectangle.Empty;

                    //-----------------------------------
                    // Padding
                    //-----------------------------------
                    left = Math.Max(0, left - 5);
                    top = Math.Max(0, top - 5);

                    right = Math.Min(bin.Cols - 1, right + 5);
                    bottom = Math.Min(bin.Rows - 1, bottom + 5);

                    return new Rectangle(
                        search.X + left,
                        search.Y + top,
                        right - left + 1,
                        bottom - top + 1);
                }
            }
        }
        private Rectangle DetectTopLeftBlock_old(Bitmap bmp)
        {
            using (Mat src = BitmapConverter.ToMat(bmp))
            using (Mat gray = new Mat())
            using (Mat thresh = new Mat())
            using (Mat morph = new Mat())
            {
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);

                // 이진화 (Otsu)
                // Cv2.Threshold(gray, thresh, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);

                // 흰색만 추출
                Cv2.Threshold(
                    gray,
                    thresh,
                    180,
                    255,
                    ThresholdTypes.Binary);

                // Morphology로 잡음 제거 및 블럭 연결
                //Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(5, 5));
                //Cv2.MorphologyEx(thresh, morph, MorphTypes.Close, kernel, iterations: 2);
                Mat kernel = Cv2.GetStructuringElement(
                                MorphShapes.Rect,
                                new OpenCvSharp.Size(35, 15));
                Cv2.MorphologyEx(
                    thresh,
                    morph,
                    MorphTypes.Close,
                    kernel,
                    iterations: 2);

                // 컨투어 검색
                OpenCvSharp.Point[][] contours;
                HierarchyIndex[] hier;
                Cv2.FindContours(morph, out contours, out hier, RetrievalModes.External, ContourApproximationModes.ApproxSimple);

                int imgW = bmp.Width;
                int imgH = bmp.Height;

                OpenCvSharp.Rect best = new OpenCvSharp.Rect();
                double bestArea = 0;

                for (int i = 0; i < contours.Length; i++)
                {
                    OpenCvSharp.Rect r = Cv2.BoundingRect(contours[i]);

                    // top-left 근처에 있는지 체크
                    if (r.X > imgW * 0.35 || r.Y > imgH * 0.30)
                        continue;
                    if (r.Width < imgW * 0.12)
                        continue;

                    if (r.Height < 40)
                        continue;

                    double area = r.Width * r.Height;

                    // 너무 작거나 너무 얇은 블럭은 무시
                    if (area < 100) continue;

                    double aspect = (double)r.Width / r.Height;
                    if (aspect < 0.3 || aspect > 3.0) continue;

                    if (area > bestArea)
                    {
                        bestArea = area;
                        best = r;
                    }
                }

                gray.SaveImage("gray.png");
                thresh.SaveImage("thresh.png");
                morph.SaveImage("morph.png");

                if (bestArea > 0)
                {
                    // 결과를 안전한 Rectangle로 변환
                    Rectangle result = new Rectangle(best.X, best.Y, best.Width, best.Height);
                    return result;
                }
            }

            return Rectangle.Empty;
        }
#endif

        private Bitmap CropRightBottom(Bitmap bmp)
        {
            Rectangle r = new Rectangle(
                (int)(bmp.Width * 0.87),
                (int)(bmp.Height * 0.9),
                (int)(bmp.Width * 0.13),
                (int)(bmp.Height * 0.09));

            return bmp.Clone(r, bmp.PixelFormat);
        }

        private Bitmap Preprocess(Bitmap src, int scale)
        {
            Bitmap bmp = new Bitmap(src.Width * scale, src.Height * scale);

            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(src, 0, 0, bmp.Width, bmp.Height);
            }

            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    Color c = bmp.GetPixel(x, y);

                    int gray = (int)(0.299 * c.R + 0.587 * c.G + 0.114 * c.B);

                    //gray = gray > 140 ? 255 : 0;

                    bmp.SetPixel(x, y, Color.FromArgb(gray, gray, gray));
                }
            }

            return bmp;
        }

        public void Dispose()
        {
            if (_ownsOcr)
            {
                _ocr?.Dispose();
            }
        }
    }

    public class FlirResult
    {
        public Dictionary<string, double> Items { get; set; } = new Dictionary<string, double>();

        public double ScaleMinTemperature { get; set; }

        public string LeftRawText { get; set; }

        public string RightRawText { get; set; }

        public void RemoveBelowScaleMinTemperature()
        {
            Items = Items
                .Where(x => x.Value >= ScaleMinTemperature)
                .ToDictionary(x => x.Key, x => x.Value);
        }
    }
}