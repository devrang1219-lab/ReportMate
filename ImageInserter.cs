using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;



//using Tesseract;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace SoborSmartReport
{
    public class ImageInserter : IDisposable
    {
        private readonly Excel.Worksheet _ws;
        private readonly List<string> _imagePaths = new List<string>(); // 삽입할 이미지들
        private readonly List<string> _tempFiles = new List<string>();  // 삭제할 임시파일들
        public int ImageCount => _imagePaths.Count;


        public ImageInserter(Excel.Worksheet ws, string filePath)
        {
            _ws = ws;


            int dpi = 300;
            int width = 2480;
            int height = 3508;

            if (Path.GetExtension(filePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                using (var pdf = PdfDocument.Load(filePath))
                {
                    for (int page = 0; page < pdf.PageCount; page++)
                    {
                        string tempPath = Path.Combine(
                            Path.GetTempPath(),
                            Guid.NewGuid() + ".png");

                        using (Image img = pdf.Render(
                            page,
                            width,
                            height,
                            dpi,
                            dpi,
                            true))
                        {
                            img.Save(tempPath, ImageFormat.Png);
                        }

                        _imagePaths.Add(tempPath);
                        _tempFiles.Add(tempPath);
                    }
                }
            }
            else
            {
                _imagePaths.Add(filePath);
            }
        }

        public Excel.Shape Insert(int page, string cellAddress, ImageInsertOptions option = null)
        {
            if (option == null)
                option = new ImageInsertOptions();

            if (page < 0 || page >= _imagePaths.Count)
                throw new ArgumentOutOfRangeException(nameof(page));

            string imagePath = GetImagePath(page, option);

            Excel.Range cell = _ws.Range[cellAddress];

            Excel.Shape picture = _ws.Shapes.AddPicture(
                imagePath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                (float)cell.Left,
                (float)cell.Top,
                -1,
                -1);

            if (option.Width.HasValue && option.Height.HasValue)
            {
                picture.LockAspectRatio = Office.MsoTriState.msoFalse;
                picture.Width = (float)option.Width.Value;
                picture.Height = (float)option.Height.Value;
            }
            else if (option.Scale.HasValue)
            {
                picture.LockAspectRatio = Office.MsoTriState.msoTrue;
                picture.ScaleWidth((float)option.Scale.Value, Office.MsoTriState.msoTrue);
                picture.ScaleHeight((float)option.Scale.Value, Office.MsoTriState.msoTrue);
            }

            return picture;
        }

        private bool NeedCrop(ImageInsertOptions option)
        {
            return option.CropLeft > 0 ||
                   option.CropTop > 0 ||
                   option.CropRight > 0 ||
                   option.CropBottom > 0;
        }

        private string GetImagePath(int page, ImageInsertOptions option)
        {
            string imagePath = _imagePaths[page];

            if (NeedCrop(option))
            {
                imagePath = CropImage(
                    imagePath,
                    option.CropLeft,
                    option.CropTop,
                    option.CropRight,
                    option.CropBottom);
            }

            return imagePath;
        }

        public Excel.Shape Insert(string cellAddress, ImageInsertOptions option)
        {
            return Insert(0, cellAddress, option);
        }

        public Excel.Shape Insert(string cellAddress, double scale)
        {
            return Insert(cellAddress, new ImageInsertOptions
            {
                Scale = scale
            });
        }

        public Excel.Shape Insert(string cellAddress, double width, double height)
        {
            return Insert(cellAddress, new ImageInsertOptions
            {
                Width = width,
                Height = height
            });
        }

        public void Dispose()
        {
            foreach (string file in _tempFiles)
            {
                try
                {
                    if (File.Exists(file))
                        File.Delete(file);
                }
                catch { }
            }
        }

        private bool IsRotatedImage(Excel.Shape shape)
        {
            double rot = shape.Rotation % 360;
            if (rot < 0) rot += 360;

            bool rotated90 =
                Math.Abs(rot - 90) < 1 ||
                Math.Abs(rot - 270) < 1;

            if (rotated90) return true;
            return false;
        }

        public void InsertFitPt(
    double left,
    double top,
    double width,
    double height,
    ImageInsertOptions option = null)
        {
            InsertFitPt(
                0,
                left,
                top,
                width,
                height,
                option);
        }

        public void InsertFitPt(
            int page,
            double left,
            double top,
            double width,
            double height,
            ImageInsertOptions option = null)
        {
            if (option == null)
                option = new ImageInsertOptions();

            if (page < 0 || page >= _imagePaths.Count)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(page));
            }

            if (width <= 0)
                throw new ArgumentOutOfRangeException(
                    nameof(width));

            if (height <= 0)
                throw new ArgumentOutOfRangeException(
                    nameof(height));

            Excel.Shape? pic = null;
            Excel.Shapes? shapes = null;

            string tempFixed = "";
            string tempResized = "";
            string imagePath = "";

            try
            {
                // =====================================================
                // 1. 이미지 경로
                // =====================================================
                imagePath =
                    GetImagePath(
                        page,
                        option);

                tempFixed =
                    FixImageOrientation(
                        imagePath);

                // =====================================================
                // 2. PT 좌표
                // =====================================================
                float imageLeft =
                    (float)left;

                float imageTop =
                    (float)top;

                float imageWidth =
                    (float)width;

                float imageHeight =
                    (float)height;

                Debug.WriteLine(
                    $"PT 영역: " +
                    $"Left={imageLeft}, " +
                    $"Top={imageTop}, " +
                    $"Width={imageWidth}, " +
                    $"Height={imageHeight}");

                // =====================================================
                // 3. 이미지 Resize
                // =====================================================
                int pixelWidth =
                    (int)Math.Ceiling(
                        imageWidth *
                        96.0 /
                        72.0 *
                        1.3);

                int pixelHeight =
                    (int)Math.Ceiling(
                        imageHeight *
                        96.0 /
                        72.0 *
                        1.3);

                tempResized =
                    ResizeImage(
                        tempFixed,
                        pixelWidth,
                        pixelHeight);

                // =====================================================
                // 4. Shapes
                // =====================================================
                shapes =
                    _ws.Shapes;

                // =====================================================
                // 5. 이미지 삽입
                // =====================================================
                pic =
                    shapes.AddPicture(
                        tempResized,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        imageLeft,
                        imageTop,
                        -1,
                        -1);

                Debug.WriteLine(
                    $"After AddPicture: " +
                    $"Left={pic.Left}, " +
                    $"Top={pic.Top}, " +
                    $"Width={pic.Width}, " +
                    $"Height={pic.Height}");

                // =====================================================
                // 6. 이미지 크기 조정
                // =====================================================
                if (option.KeepAspectRatio)
                {
                    pic.LockAspectRatio =
                        Office.MsoTriState.msoTrue;

                    double scaleX =
                        imageWidth /
                        pic.Width;

                    double scaleY =
                        imageHeight /
                        pic.Height;

                    double scale =
                        Math.Min(
                            scaleX,
                            scaleY);

                    pic.ScaleWidth(
                        (float)scale,
                        Office.MsoTriState.msoTrue);

                    pic.ScaleHeight(
                        (float)scale,
                        Office.MsoTriState.msoTrue);

                    // 가운데 정렬
                    pic.Left =
                        imageLeft +
                        (imageWidth - pic.Width) / 2f +
                        option.GapLeft;

                    pic.Top =
                        imageTop +
                        (imageHeight - pic.Height) / 2f +
                        option.GapTop;

                    pic.Width -=
                        option.GapLeft +
                        option.GapRight;

                    pic.Height -=
                        option.GapTop +
                        option.GapBottom;
                }
                else
                {
                    pic.LockAspectRatio =
                        Office.MsoTriState.msoFalse;

                    pic.Left =
                        imageLeft +
                        option.GapLeft;

                    pic.Top =
                        imageTop +
                        option.GapTop;

                    pic.Width =
                        imageWidth -
                        option.GapLeft -
                        option.GapRight;

                    pic.Height =
                        imageHeight -
                        option.GapTop -
                        option.GapBottom;

                    // =================================================
                    // 회전 이미지
                    // =================================================
                    if (IsRotatedImage(pic))
                    {
                        Debug.WriteLine(
                            $"Rotated Image: " +
                            $"{pic.Name}, " +
                            $"Rotation: {pic.Rotation}");

                        pic.Width =
                            imageHeight -
                            option.GapLeft -
                            option.GapRight;

                        pic.Height =
                            imageWidth -
                            option.GapTop -
                            option.GapBottom;
                    }
                }

                Debug.WriteLine(
                    $"PT 이미지: " +
                    $"Left={pic.Left}, " +
                    $"Top={pic.Top}, " +
                    $"Width={pic.Width}, " +
                    $"Height={pic.Height}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"InsertFitPt 오류: {ex.Message}");

                if (pic != null)
                {
                    try
                    {
                        pic.Delete();
                    }
                    catch
                    {
                    }
                }

                throw;
            }
            finally
            {
                // =====================================================
                // Shape
                // =====================================================
                if (pic != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(pic);
                    }
                    catch
                    {
                    }

                    pic = null;
                }

                // =====================================================
                // Shapes
                // =====================================================
                if (shapes != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(shapes);
                    }
                    catch
                    {
                    }

                    shapes = null;
                }

                // =====================================================
                // 임시 파일
                // =====================================================
                DeleteTempFile(
                    tempFixed,
                    imagePath);

                DeleteTempFile(
                    tempResized,
                    imagePath);
            }
        }

        public void InsertFit(
            string cellFrom,
            string cellTo,
            ImageInsertOptions option = null)
        {
            InsertFit(0, cellFrom, cellTo, option);
        }

        public string ResizeImage(string path, int maxWidth, int maxHeight)
        {
            using (Image src = Image.FromFile(path))
            {
                int width = src.Width;
                int height = src.Height;

                double ratio = Math.Min(
                    (double)maxWidth / width,
                    (double)maxHeight / height);

                // 원본보다 크게 만들지 않음
                if (ratio >= 1.0)
                    return path;

                int newWidth = (int)Math.Round(width * ratio);
                int newHeight = (int)Math.Round(height * ratio);

                using (Bitmap bmp = new Bitmap(newWidth, newHeight))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                        g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                        g.DrawImage(src, 0, 0, newWidth, newHeight);
                    }

                    string newPath = Path.Combine(
                        Path.GetDirectoryName(path),
                        "resize_" + Path.GetFileName(path));

                    // JPEG 품질 설정 (90 정도 추천)
                    ImageCodecInfo jpgEncoder = ImageCodecInfo.GetImageEncoders()
                        .First(c => c.FormatID == ImageFormat.Jpeg.Guid);

                    EncoderParameters ep = new EncoderParameters(1);
                    ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 90L);

                    bmp.Save(newPath, jpgEncoder, ep);

                    return newPath;
                }
            }
        }

        public static string FixImageOrientation(string path)
        {
            using (Image img = Image.FromFile(path))
            {
                if (img == null || img.PropertyIdList == null || img.PropertyIdList.Length == 0)
                    return path;

                const int OrientationId = 0x0112;

                if (!img.PropertyIdList.Contains(OrientationId))
                    return path;

                var prop = img.GetPropertyItem(OrientationId);

                int orientation = BitConverter.ToUInt16(prop.Value, 0);

                RotateFlipType rotateFlip = RotateFlipType.RotateNoneFlipNone;

                switch (orientation)
                {
                    case 2:
                        rotateFlip = RotateFlipType.RotateNoneFlipX;
                        break;

                    case 3:
                        rotateFlip = RotateFlipType.Rotate180FlipNone;
                        break;

                    case 4:
                        rotateFlip = RotateFlipType.Rotate180FlipX;
                        break;

                    case 5:
                        rotateFlip = RotateFlipType.Rotate90FlipX;
                        break;

                    case 6:
                        rotateFlip = RotateFlipType.Rotate90FlipNone;
                        break;

                    case 7:
                        rotateFlip = RotateFlipType.Rotate270FlipX;
                        break;

                    case 8:
                        rotateFlip = RotateFlipType.Rotate270FlipNone;
                        break;
                }

                if (rotateFlip != RotateFlipType.RotateNoneFlipNone)
                {
                    img.RotateFlip(rotateFlip);

                    // EXIF 방향 제거
                    img.RemovePropertyItem(OrientationId);
                }

                string newPath = Path.Combine(
                    Path.GetDirectoryName(path),
                    "fixed_" + Path.GetFileName(path));

                img.Save(newPath, ImageFormat.Jpeg);

                return newPath;
            }
        }

        private void DeleteTempFile(string path, string originalPath)
        {
            if (string.IsNullOrEmpty(path))
                return;

            if (string.Equals(
                path,
                originalPath,
                StringComparison.OrdinalIgnoreCase))
                return;

            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"임시 이미지 삭제 실패: {path}, {ex.Message}");
            }
        }

        public void InsertFit(
            int page,
            string cellFrom,
            string cellTo,
            ImageInsertOptions option = null)
        {
            if (option == null)
                option = new ImageInsertOptions();

            if (page < 0 || page >= _imagePaths.Count)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(page));
            }

            Excel.Shape? pic = null;

            Excel.Shapes? shapes = null;

            Excel.Range? area = null;
            Excel.Range? fromCell = null;
            Excel.Range? toCell = null;

            string tempFixed = "";
            string tempResized = "";
            string imagePath = "";

            try
            {
                // =====================================================
                // 1. 이미지 경로
                // =====================================================
                imagePath =
                    GetImagePath(
                        page,
                        option);

                tempFixed =
                    FixImageOrientation(
                        imagePath);

                // =====================================================
                // 2. 셀 영역
                // =====================================================
                area =
                    _ws.Range[
                        cellFrom,
                        cellTo];

                fromCell =
                    _ws.Range[cellFrom];

                toCell =
                    _ws.Range[cellTo];

                Debug.WriteLine(
                    $"cell from: {cellFrom}, " +
                    $"to: {cellTo}");

                Debug.WriteLine(
                    $"image name: {imagePath}");

                Debug.WriteLine(
                    $"area: {area.Left}, " +
                    $"{area.Top}, " +
                    $"{area.Width}, " +
                    $"{area.Height}");

                // =====================================================
                // 3. 이미지 위치 / 크기
                // =====================================================
                float left = Convert.ToSingle(fromCell.Left);

                float top = Convert.ToSingle(fromCell.Top);

                float width = Convert.ToSingle(
                         (double)toCell.Left +
                         (double)toCell.Width -
                         left);

                float height = Convert.ToSingle(
                        (double)toCell.Top +
                        (double)toCell.Height -
                        top);

                // =====================================================
                // 4. 테두리 보정
                // =====================================================
                float borderLeftGap =
                    GetBorderGap(
                        fromCell,
                        Excel.XlBordersIndex.xlEdgeLeft);

                float borderTopGap =
                    GetBorderGap(
                        fromCell,
                        Excel.XlBordersIndex.xlEdgeTop);

                // =====================================================
                // 5. 이미지 Resize
                // =====================================================
                int pixelWidth =
                    (int)Math.Ceiling(
                        width *
                        96.0 /
                        72.0 *
                        1.3);

                int pixelHeight =
                    (int)Math.Ceiling(
                        height *
                        96.0 /
                        72.0 *
                        1.3);

                tempResized =
                    ResizeImage(
                        tempFixed,
                        pixelWidth,
                        pixelHeight);

                // =====================================================
                // 6. Shapes 컬렉션 확보
                // =====================================================
                shapes =
                    _ws.Shapes;

                // =====================================================
                // 7. 이미지 삽입
                // =====================================================
                pic =
                    shapes.AddPicture(
                        tempResized,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        left,
                        top,
                        -1,
                        -1);

                Debug.WriteLine(
                    "After AddPicture");

                Debug.WriteLine(
                    $"Left={pic.Left}");

                Debug.WriteLine(
                    $"Top={pic.Top}");

                Debug.WriteLine(
                    $"Width={pic.Width}");

                Debug.WriteLine(
                    $"Height={pic.Height}");

                Debug.WriteLine(
                    $"Rotation={pic.Rotation}");

                // =====================================================
                // 8. 이미지 크기 조정
                // =====================================================
                if (option.KeepAspectRatio)
                {
                    pic.LockAspectRatio =
                        Office.MsoTriState.msoTrue;

                    double scaleX =
                        width /
                        pic.Width;

                    double scaleY =
                        height /
                        pic.Height;

                    double scale =
                        Math.Min(
                            scaleX,
                            scaleY);

                    pic.ScaleWidth(
                        (float)scale,
                        Office.MsoTriState.msoTrue);

                    pic.ScaleHeight(
                        (float)scale,
                        Office.MsoTriState.msoTrue);

                    // 가운데 정렬
                    pic.Left =
                        left +
                        (width - pic.Width) / 2f +
                        option.GapLeft;

                    pic.Top =
                        top +
                        (height - pic.Height) / 2f +
                        option.GapTop;

                    pic.Width -=
                        option.GapLeft +
                        option.GapRight;

                    pic.Height -=
                        option.GapTop +
                        option.GapBottom;
                }
                else
                {
                    pic.LockAspectRatio =
                        Office.MsoTriState.msoFalse;

                    pic.Left =
                        left +
                        option.GapLeft;

                    pic.Top =
                        top +
                        option.GapTop;

                    pic.Width =
                        width -
                        option.GapLeft -
                        option.GapRight;

                    pic.Height =
                        height -
                        option.GapTop -
                        option.GapBottom;

                    // =================================================
                    // 회전 이미지
                    // =================================================
                    if (IsRotatedImage(pic))
                    {
                        Debug.WriteLine(
                            $"Rotated Image: " +
                            $"{pic.Name}, " +
                            $"Rotation: {pic.Rotation}");

                        pic.Width =
                            height -
                            option.GapLeft -
                            option.GapRight;

                        pic.Height =
                            width -
                            option.GapTop -
                            option.GapBottom;
                    }
                }

                // =====================================================
                // 9. 테두리 보정
                // =====================================================
                pic.Left +=
                    borderLeftGap;

                pic.Top +=
                    borderTopGap;

                pic.Width -=
                    borderLeftGap;

                pic.Height -=
                    borderTopGap;

                Debug.WriteLine(
                    $"pic left: {pic.Left}, " +
                    $"top: {pic.Top}, " +
                    $"width: {pic.Width}, " +
                    $"height: {pic.Height}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"InsertFit 오류: {ex.Message}");

                // 이미지가 생성된 상태에서 오류가 난 경우
                if (pic != null)
                {
                    try
                    {
                        pic.Delete();
                    }
                    catch
                    {
                    }
                }

                throw;
            }
            finally
            {
                // =====================================================
                // Shape
                // =====================================================
                if (pic != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(
                            pic);
                    }
                    catch
                    {
                    }

                    pic = null;
                }

                // =====================================================
                // Shapes
                // =====================================================
                if (shapes != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(
                            shapes);
                    }
                    catch
                    {
                    }

                    shapes = null;
                }

                // =====================================================
                // Range
                // =====================================================
                if (area != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(
                            area);
                    }
                    catch
                    {
                    }

                    area = null;
                }

                if (fromCell != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(
                            fromCell);
                    }
                    catch
                    {
                    }

                    fromCell = null;
                }

                if (toCell != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(
                            toCell);
                    }
                    catch
                    {
                    }

                    toCell = null;
                }

                // =====================================================
                // 임시 파일
                // =====================================================
                DeleteTempFile(
                    tempFixed,
                    imagePath);

                DeleteTempFile(
                    tempResized,
                    imagePath);
            }
        }

        private float GetBorderGap(Excel.Range cell, Excel.XlBordersIndex borderIndex)
        {
            Excel.Border border = null;

            try
            {
                border = cell.Borders[borderIndex];

                if (border == null)
                    return 0f;

                // 선 없음
                if ((Excel.XlLineStyle)border.LineStyle == Excel.XlLineStyle.xlLineStyleNone)
                    return 0f;

                var lineStyle = (Excel.XlLineStyle)border.LineStyle;
                var weight = (Excel.XlBorderWeight)border.Weight;

                // Excel Shape 좌표 단위는 pt
                float gap = 0f;

                // 먼저 굵기 기준
                switch (weight)
                {
                    case Excel.XlBorderWeight.xlHairline:
                        gap = 0.25f;
                        break;

                    case Excel.XlBorderWeight.xlThin:
                        gap = 0.75f;
                        break;

                    case Excel.XlBorderWeight.xlMedium:
                        gap = 1.5f;
                        break;

                    case Excel.XlBorderWeight.xlThick:
                        var offset = (borderIndex == Excel.XlBordersIndex.xlEdgeLeft) ? 0.75f : 0f;
                        gap = 2.25f - offset;
                        break;
                }

                // 이중선은 실제로 안쪽 공간을 더 차지하므로 추가 보정
                if (lineStyle == Excel.XlLineStyle.xlDouble)
                {
                    gap = Math.Max(gap, 2.5f);
                }

                return gap;
            }
            catch
            {
                return 0f;
            }
        }


        private string CropWhiteMargin(string imagePath, byte threshold = 250)
        {
            using (Bitmap bmp = new Bitmap(imagePath))
            {
                int left = bmp.Width;
                int right = 0;
                int top = bmp.Height;
                int bottom = 0;

                for (int y = 0; y < bmp.Height; y++)
                {
                    for (int x = 0; x < bmp.Width; x++)
                    {
                        Color c = bmp.GetPixel(x, y);

                        // 흰색이 아닌 픽셀
                        if (c.R < threshold || c.G < threshold || c.B < threshold)
                        {
                            if (x < left) left = x;
                            if (x > right) right = x;
                            if (y < top) top = y;
                            if (y > bottom) bottom = y;
                        }
                    }
                }

                if (left >= right || top >= bottom)
                    return imagePath;

                System.Drawing.Rectangle crop = System.Drawing.Rectangle.FromLTRB(left, top, right + 1, bottom + 1);

                using (Bitmap cropped = bmp.Clone(crop, bmp.PixelFormat))
                {
                    string temp = Path.Combine(
                        Path.GetTempPath(),
                        Guid.NewGuid() + ".png");

                    cropped.Save(temp, ImageFormat.Png);

                    return temp;
                }
            }
        }

        private string CropImage(
            string imagePath,
            int left,
            int top,
            int right,
            int bottom)
        {
            using (Bitmap source = new Bitmap(imagePath))
            {
                int width = source.Width - left - right;
                int height = source.Height - top - bottom;

                if (width <= 0 || height <= 0)
                    throw new ArgumentException("Crop 크기가 이미지보다 큽니다.");

                System.Drawing.Rectangle cropRect = new System.Drawing.Rectangle(left, top, width, height);

                using (Bitmap cropped = new Bitmap(width, height))
                using (Graphics g = Graphics.FromImage(cropped))
                {
                    g.DrawImage(
                        source,
                        new System.Drawing.Rectangle(0, 0, width, height),
                        cropRect,
                        GraphicsUnit.Pixel);

                    string tempPath = Path.Combine(
                        Path.GetTempPath(),
                        Guid.NewGuid().ToString() + ".png");

                    cropped.Save(tempPath, ImageFormat.Png);

                    return tempPath;
                }
            }
        }
        
    }



    public class ImageInsertOptions
    {
        // 위치
        public string CellAddress { get; set; }



        // 크기
        public double? Scale { get; set; }
        public double? Width { get; set; }
        public double? Height { get; set; }


        // 비율 유지
        public bool KeepAspectRatio { get; set; } = false;

        // Crop
        public int CropLeft { get; set; }
        public int CropTop { get; set; }
        public int CropRight { get; set; }
        public int CropBottom { get; set; }

        // Gap (Point)
        public float GapLeft { get; set; } = 2.8f;
        public float GapTop { get; set; } = 2.8f;
        public float GapRight { get; set; } = 2.8f;
        public float GapBottom { get; set; } = 2.8f;
    }
}