
using System;
using System.Drawing;
using System.IO;
using Tesseract;

namespace SoborSmartReport.Ocr
{
    public sealed class OcrReader : IDisposable
    {
        private readonly TesseractEngine _engine;
        private readonly object _engineLock = new object();
        private bool _disposed;

        /// <summary>
        /// 기본 tessdata 경로:
        /// 실행 폴더/Ocr/tessdata
        /// </summary>
        public OcrReader(
            string whitelist = "0123456789.-:%/ ")
            : this(
                Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "Ocr",
                    "tessdata"),
                whitelist)
        {
        }

        /// <summary>
        /// tessdata 경로를 지정하여 OCR 엔진을 생성한다.
        /// eng.traineddata 파일이 지정 폴더에 있어야 한다.
        /// </summary>
        public OcrReader(
            string tessDataPath,
            string whitelist)
        {
            if (string.IsNullOrWhiteSpace(tessDataPath))
            {
                throw new ArgumentException(
                    "Tesseract tessdata 경로가 비어 있습니다.",
                    nameof(tessDataPath));
            }

            if (!Directory.Exists(tessDataPath))
            {
                throw new DirectoryNotFoundException(
                    $"Tesseract 데이터 폴더를 찾을 수 없습니다: {tessDataPath}");
            }

            string trainedDataFile =
                Path.Combine(tessDataPath, "eng.traineddata");

            if (!File.Exists(trainedDataFile))
            {
                throw new FileNotFoundException(
                    "eng.traineddata 파일을 찾을 수 없습니다.",
                    trainedDataFile);
            }

            _engine = new TesseractEngine(
                tessDataPath,
                "eng",
                EngineMode.Default);

            try
            {
                SetWhitelist(whitelist);

                _engine.SetVariable(
                    "preserve_interword_spaces",
                    "1");
            }
            catch
            {
                _engine.Dispose();
                throw;
            }
        }

        private void SetWhitelist(string whitelist)
        {
            if (!string.IsNullOrEmpty(whitelist))
            {
                _engine.SetVariable(
                    "tessedit_char_whitelist",
                    whitelist);
            }
        }

        /// <summary>
        /// Bitmap 이미지에서 OCR을 수행한다.
        /// 호출자가 Bitmap을 생성했다면 호출자가 Bitmap을 해제한다.
        /// </summary>

        public OcrReadResult Read(
            Bitmap bitmap,
            PageSegMode pageSegMode = PageSegMode.SingleLine)
        {
            if (bitmap == null)
                throw new ArgumentNullException(nameof(bitmap));

            ThrowIfDisposed();

            lock (_engineLock)
            {
                ThrowIfDisposed();

                // Bitmap을 PNG 바이트 배열로 변환
                using (var stream = new MemoryStream())
                {
                    bitmap.Save(
                        stream,
                        System.Drawing.Imaging.ImageFormat.Png);

                    byte[] imageBytes = stream.ToArray();

                    // PixConverter 대신 메모리에서 Pix 생성
                    using (Pix pix = Pix.LoadFromMemory(imageBytes))
                    using (Page page = _engine.Process(pix, pageSegMode))
                    {
                        string text = page.GetText() ?? "";

                        return new OcrReadResult
                        {
                            Text = text.Trim(),
                            Confidence = page.GetMeanConfidence()
                        };
                    }
                }
            }
        }



        /// <summary>
        /// 이미지 파일 경로에서 OCR을 수행한다.
        /// 지원 형식은 설치된 이미지 디코더에 따라 달라질 수 있다.
        /// </summary>
        public OcrReadResult ReadFile(
            string imagePath,
            PageSegMode pageSegMode = PageSegMode.SingleLine)
        {
            if (string.IsNullOrWhiteSpace(imagePath))
            {
                throw new ArgumentException(
                    "이미지 경로가 비어 있습니다.",
                    nameof(imagePath));
            }

            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException(
                    "OCR 대상 이미지 파일을 찾을 수 없습니다.",
                    imagePath);
            }

            ThrowIfDisposed();

            using (Bitmap bitmap = new Bitmap(imagePath))
            {
                return Read(bitmap, pageSegMode);
            }
        }

        /// <summary>
        /// OCR 결과를 최소 신뢰도 기준으로 판정한다.
        /// minConfidence 범위는 0.0 ~ 1.0이다.
        /// </summary>
        public bool TryRead(
            Bitmap bitmap,
            out OcrReadResult result,
            float minConfidence = 0.5f,
            PageSegMode pageSegMode = PageSegMode.SingleLine)
        {
            if (minConfidence < 0f || minConfidence > 1f)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(minConfidence),
                    "신뢰도는 0.0부터 1.0 사이여야 합니다.");
            }

            result = Read(bitmap, pageSegMode);

            return result.HasText &&
                   result.Confidence >= minConfidence;
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(OcrReader));
            }
        }

        public void Dispose()
        {
            lock (_engineLock)
            {
                if (_disposed)
                    return;

                _disposed = true;
                _engine.Dispose();
            }

            GC.SuppressFinalize(this);
        }
    }

    public sealed class OcrReadResult
    {
        public string Text { get; set; } = "";

        /// <summary>
        /// Tesseract 평균 신뢰도.
        /// 일반적으로 0.0 ~ 1.0 범위.
        /// </summary>
        public float Confidence { get; set; }

        public bool HasText =>
            !string.IsNullOrWhiteSpace(Text);

        public bool IsHighConfidence(float threshold = 0.7f) =>
            HasText && Confidence >= threshold;

        public override string ToString()
        {
            return $"{Text} ({Confidence:P1})";
        }
    }
}
