using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#if OPENCV
using OpenCvSharp;
using OpenCvSharp.Extensions;
#endif

namespace SoborSmartReport.Ocr
{
    public static class OcrImageProcessor
    {
        /// <summary>
        /// 이미지 크기 대비 비율로 Crop.
        /// x/y/width/height = 0.0 ~ 1.0
        /// </summary>
        public static Bitmap CropRatio(
            Bitmap source,
            double x,
            double y,
            double width,
            double height)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            Rectangle rect = GetRatioRectangle(
                source,
                x,
                y,
                width,
                height);

            return source.Clone(
                rect,
                source.PixelFormat);
        }

        public static Rectangle GetRatioRectangle(
            Bitmap source,
            double x,
            double y,
            double width,
            double height)
        {
            int px = (int)Math.Round(
                source.Width * x);

            int py = (int)Math.Round(
                source.Height * y);

            int pw = (int)Math.Round(
                source.Width * width);

            int ph = (int)Math.Round(
                source.Height * height);

            px = Math.Max(0, px);
            py = Math.Max(0, py);

            pw = Math.Min(
                pw,
                source.Width - px);

            ph = Math.Min(
                ph,
                source.Height - py);

            if (pw <= 0 || ph <= 0)
            {
                throw new ArgumentException(
                    $"잘못된 ROI입니다. " +
                    $"x={x}, y={y}, width={width}, height={height}");
            }

            return new Rectangle(
                px,
                py,
                pw,
                ph);
        }

#if OPENCV

        public static Bitmap Preprocess(
            Bitmap source,
            double scale = 3.0,
            bool grayscale = true,
            string threshold = "otsu")
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            using (Mat src =
                BitmapConverter.ToMat(source))
            using (Mat work = new Mat())
            using (Mat resized = new Mat())
            using (Mat result = new Mat())
            {
                //--------------------------------------
                // Grayscale
                //--------------------------------------

                if (grayscale)
                {
                    if (src.Channels() == 4)
                    {
                        Cv2.CvtColor(
                            src,
                            work,
                            ColorConversionCodes.BGRA2GRAY);
                    }
                    else if (src.Channels() == 3)
                    {
                        Cv2.CvtColor(
                            src,
                            work,
                            ColorConversionCodes.BGR2GRAY);
                    }
                    else
                    {
                        src.CopyTo(work);
                    }
                }
                else
                {
                    src.CopyTo(work);
                }

                //--------------------------------------
                // 확대
                //--------------------------------------

                if (scale > 0 &&
                    Math.Abs(scale - 1.0) > 0.001)
                {
                    Cv2.Resize(
                        work,
                        resized,
                        Size.Zero,
                        scale,
                        scale,
                        InterpolationFlags.Cubic);
                }
                else
                {
                    work.CopyTo(resized);
                }

                //--------------------------------------
                // Threshold
                //--------------------------------------

                if (string.Equals(
                    threshold,
                    "otsu",
                    StringComparison.OrdinalIgnoreCase))
                {
                    // Otsu는 단일 채널 필요
                    if (resized.Channels() == 1)
                    {
                        Cv2.Threshold(
                            resized,
                            result,
                            0,
                            255,
                            ThresholdTypes.Binary |
                            ThresholdTypes.Otsu);
                    }
                    else
                    {
                        using (Mat gray = new Mat())
                        {
                            Cv2.CvtColor(
                                resized,
                                gray,
                                ColorConversionCodes.BGR2GRAY);

                            Cv2.Threshold(
                                gray,
                                result,
                                0,
                                255,
                                ThresholdTypes.Binary |
                                ThresholdTypes.Otsu);
                        }
                    }
                }
                else
                {
                    resized.CopyTo(result);
                }

                return BitmapConverter.ToBitmap(result);
            }
        }

#endif
    }
}
