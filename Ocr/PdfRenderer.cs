using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoborSmartReport.Ocr
{
    public static class PdfImageRenderer
    {
        public static Bitmap RenderPage(
            string pdfPath,
            int page = 0,
            int dpi = 400)
        {
            using (PdfDocument pdf =
                PdfDocument.Load(pdfPath))
            {
                if (page < 0 ||
                    page >= pdf.PageCount)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(page));
                }

                // PDF 실제 페이지 크기(point)
                SizeF pageSize =
                    pdf.PageSizes[page];

                // PDF point = 1/72 inch
                int width = (int)Math.Round(
                    pageSize.Width / 72.0 * dpi);

                int height = (int)Math.Round(
                    pageSize.Height / 72.0 * dpi);

                using (Image image = pdf.Render(
                    page,
                    width,
                    height,
                    dpi,
                    dpi,
                    true))
                {
                    // PdfDocument이 Dispose되어도 사용할 수 있도록 복사
                    return new Bitmap(image);
                }
            }
        }
    }
}
