using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SoborSmartReport.Comm;
using SoborSmartReport.Ocr;
using Excel = Microsoft.Office.Interop.Excel;

namespace SoborSmartReport
{
    public class Report
    {
        public bool isAnnual { get; set; } = false;
        public bool isHalfYear { get; set; } = false;
        public bool isOnlyAnnual { get; set; } = false;
        public bool isUpperOfHalfYear { get; set; } = false;
        public int nQuater { get; set; } = 1;
        public int nYear { get; set; } = 2026;
        public int nMonth { get; set; } = 7;
        public int nDay { get; set; } = 1;
        public int totalCount { get; set; } = 0;
        public int quaterCount { get; set; } = 0;
        public string strSite { get; set; }
        public string strInspector { get; set; }
        public string xlsFilePath { get; set; }

        private string xlsSamplePath { get; set; }

        public Action<string, string> LogHandler { get; set; }

        public static Report? ParseReport(string filePath, string samplePath, SoborLog soborLog)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);

            string[] parts = fileName.Split('_');

            if (parts.Length < 2) return null;

            string date = parts[parts.Length-1].Trim();

            if (date.Length < 6) return null;

            Report report = new Report();
            report.xlsFilePath = filePath;
            report.xlsSamplePath = samplePath;

            report.strSite = parts[0];

            string reportName = parts[1];

            // 연차 여부
            report.isAnnual = reportName.Contains("연차");

            // 연도
            Match yearMatch = Regex.Match(reportName, @"(\d{2})년");
            if (yearMatch.Success)
                report.nYear = 2000 + int.Parse(yearMatch.Groups[1].Value);

            // 분기
            //Match quarterMatch = Regex.Match(reportName, @"(\d)분기");
            //if (quarterMatch.Success)
            //    report.nQuater = int.Parse(quarterMatch.Groups[1].Value);

            // 날짜
            if (parts.Length >= 3 && parts[2].Trim().Length == 6)
            {
                report.nYear = 2000 + int.Parse(date.Substring(0, 2));
                report.nMonth = int.Parse(date.Substring(2, 2));
                report.nDay = int.Parse(date.Substring(4, 2));
            }
            else
            {
                return null;
            }

            // 상반기 여부
            report.isUpperOfHalfYear = report.nQuater <= 2;

            report.quaterCount = report.GetQuarterCount(filePath);

            return report;
        }

        public Excel.Worksheet? GetWorksheetByName(
            Excel.Workbook wb,
            string sheetName)
        {
            Excel.Sheets? sheets = null;

            try
            {
                sheets = wb.Worksheets;

                int count = sheets.Count;

                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet? sheet = null;

                    try
                    {
                        sheet = (Excel.Worksheet)sheets[i];

                        bool matched =
                            sheet.Name.Trim()
                                .IndexOf(
                                    sheetName,
                                    StringComparison.OrdinalIgnoreCase) >= 0;

                        if (matched)
                        {
                            // 소유권을 호출한 쪽으로 넘김
                            Excel.Worksheet result = sheet;
                            sheet = null;

                            return result;
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                            Marshal.ReleaseComObject(sheet);
                    }
                }

                return null;
            }
            finally
            {
                if (sheets != null)
                    Marshal.ReleaseComObject(sheets);
            }
        }

        private int GetMonthFromFileName(string filePath)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);

            Match m = Regex.Match(fileName, @"(\d{6})$");

            if (!m.Success)
                throw new Exception("파일명에서 날짜를 찾을 수 없습니다.");

            string yymmdd = m.Groups[1].Value;

            return int.Parse(yymmdd.Substring(2, 2)); // MM
        }

        public int GetQuarterCount()
        {
            var quarterCount = 1;
            quarterCount = this.nMonth / 3 + 1;
            return quarterCount;
        }

        public int GetQuarterCount(string filePath)
        {
            totalCount = 0;
            quaterCount = 0;


            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                //MessageBox.Show("페이지 번호를 매길 엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            }

            int targetMonth = GetMonthFromFileName(filePath);
            Excel.Application? xlApp = null;
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };

                wb = ExcelComHelper.OpenWorkbook(xlApp, filePath, false);
                ws = ExcelComHelper.GetWorksheet(wb, "연계획");

                if (ws == null)
                {
                    ws = ExcelComHelper.GetWorksheet(wb, "Sheet7");
                }

                if (ws == null)
                {
                    throw new Exception("연계획 시트를 찾을 수 없습니다.");
                }

                // D열 = 4, 1월이라고 가정
                for (int month = 1; month <= 12; month++)
                {
                    int col = 3 + month; // D=4

                    string value = ExcelComHelper.GetCellText(ws, 24, col);

                    if (value == "●")
                    {
                        if (month <= targetMonth)
                            quaterCount++;
                        totalCount++;
                    }
                }

                // 절연 점검 포함은 연차, 절연점검 미포함, 접지저항측정은 반기
                string val1 = ExcelComHelper.GetCellText(ws, 12, 3 + targetMonth);
                string val2 = ExcelComHelper.GetCellText(ws, 10, 3 + targetMonth);

                if (val1 == "●" && val2 != "●")
                {
                    isHalfYear = true;
                }

                return quaterCount;
            }
            catch (Exception ex)
            {
                AddLog($"분기 수 계산 중 오류 발생: {ex.Message}");
                return 0;
            }
            finally
            {
                if (totalCount == 0)
                {
                    //MessageBox.Show(
                    //    "연계획 시트에서 분기 수를 찾을 수 없습니다.",
                    //    "오류",
                    //    MessageBoxButtons.OK,
                    //    MessageBoxIcon.Warning);

                    isOnlyAnnual = false;
                }
                else
                {
                    isOnlyAnnual = quaterCount == 1;
                }

                nQuater = quaterCount;

                ExcelComHelper.Release(ws);
                ws = null;

                ExcelComHelper.CloseWorkbook(
                    ref wb,
                    false);

                ExcelComHelper.QuitApplication(
                    ref xlApp);

                ExcelComHelper.Cleanup();
            }


        }
        public void CopySheetToXlsxAndProcess(
            string sheetName,
            float gapLeft,
            float gapTop,
            float gapRight,
            float gapBottom)
        {
            Excel.Application? xlApp = null;
            Excel.Workbooks? workbooks = null;
            Excel.Workbook? srcWb = null;
            Excel.Workbook? newWb = null;

            Excel.Worksheet? srcWs = null;
            Excel.Worksheet? newWs = null;

            Excel.Sheets? newSheets = null;

            try
            {
                // =====================================================
                // 1. Excel Application
                // =====================================================
                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                // =====================================================
                // 2. Workbooks COM 객체를 명시적으로 보관
                // =====================================================
                workbooks = xlApp.Workbooks;


                // =====================================================
                // 2. 원본 Workbook 열기
                // =====================================================
                srcWb = workbooks.Open(
                    xlsFilePath,
                    ReadOnly: false);


                // =====================================================
                // 3. 원본 Sheet 가져오기
                // =====================================================
                srcWs = GetWorksheetByName(
                    srcWb,
                    sheetName);

                if (srcWs == null)
                {
                    throw new InvalidOperationException(
                        $"원본 파일에 '{sheetName}' 시트가 없습니다.");
                }


                // =====================================================
                // 4. Sheet를 새로운 Workbook으로 복사
                // =====================================================
                srcWs.Copy(
                    Type.Missing,
                    Type.Missing);


                // Copy 후 ActiveWorkbook = 새 Workbook
                newWb = xlApp.ActiveWorkbook;

                if (newWb == null)
                {
                    throw new InvalidOperationException(
                        "시트 복사 후 새 Workbook을 가져오지 못했습니다.");
                }


                //// =====================================================
                //// 5. 저장 경로
                //// =====================================================
                string? destFolder =
                    Path.GetDirectoryName(xlsFilePath);

                string destFileName =
                    sheetName + ".xlsx";

                string destPath =
                    Path.Combine(
                        destFolder,
                        destFileName);


                //// =====================================================
                //// 6. 기존 파일 삭제
                //// =====================================================
                if (File.Exists(destPath))
                {
                    File.Delete(destPath);
                }


                //// =====================================================
                //// 7. 새 Workbook을 XLSX로 저장
                //// =====================================================
                newWb.SaveAs(
                    destPath,
                    Excel.XlFileFormat.xlOpenXMLWorkbook);

                // =====================================================
                // 10. Sheets COM 객체를 별도로 관리
                // =====================================================
                newSheets = newWb.Sheets;


                //// =====================================================
                //// 8. 새 Workbook의 첫 번째 Sheet 가져오기
                //// =====================================================
                newWs =
                 (Excel.Worksheet)newSheets[1];


                if (newWs == null)
                {
                    throw new InvalidOperationException(
                        "새 Workbook의 Sheet를 가져오지 못했습니다.");
                }


                // =====================================================
                // 9. 새 Sheet의 이미지 위치/크기 조정
                // =====================================================
                SnapImageMergedCell(
                    newWs,
                    sheetName,
                    gapLeft,
                    gapTop,
                    gapRight,
                    gapBottom);
                SnapImageMergedCell(
                    srcWs,
                    sheetName,
                    gapLeft,
                    gapTop,
                    gapRight,
                    gapBottom);


                // =====================================================
                // 10. 처리 결과 저장
                // =====================================================

                srcWb.Save();
                newWb.Save();
            }
            catch (Exception ex)
            {
                AddLog(
                    $"시트 복사 및 이미지 처리 중 오류 발생: {ex.Message}");

                //MessageBox.Show(
                //    $"시트 복사 및 처리 중 오류가 발생했습니다.\r\n\r\n{ex.Message}",
                //    "오류",
                //    MessageBoxButtons.OK,
                //    MessageBoxIcon.Error);
            }
            finally
            {
                // =====================================================
                // 11. Worksheet COM 해제
                // =====================================================
                if (newWs != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(newWs);
                    }
                    catch
                    {
                    }

                    newWs = null;
                }

                if (srcWs != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(srcWs);
                    }
                    catch
                    {
                    }

                    srcWs = null;
                }

                if (newSheets != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(newSheets);
                    }
                    catch
                    {
                    }

                    newSheets = null;
                }



                // =====================================================
                // 12. 새 Workbook 저장 및 종료
                // =====================================================
                if (newWb != null)
                {
                    try
                    {
                        newWb.Save();
                    }
                    catch
                    {
                    }

                    try
                    {
                        newWb.Close(
                            SaveChanges: false);
                    }
                    catch
                    {
                    }

                    try
                    {
                        Marshal.ReleaseComObject(newWb);
                    }
                    catch
                    {
                    }

                    newWb = null;
                }


                // =====================================================
                // 13. 원본 Workbook 종료
                // =====================================================
                if (srcWb != null)
                {
                    try
                    {
                        srcWb.Close(
                            SaveChanges: false);
                    }
                    catch
                    {
                    }

                    try
                    {
                        Marshal.ReleaseComObject(srcWb);
                    }
                    catch
                    {
                    }

                    srcWb = null;
                }

                // =====================================================
                // 2. Workbooks 컬렉션 해제
                // =====================================================
                if (workbooks != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(workbooks);
                    }
                    catch
                    {
                    }

                    workbooks = null;
                }


                // =====================================================
                // 14. Excel 종료
                // =====================================================
                if (xlApp != null)
                {
                    try
                    {
                        xlApp.Quit();
                    }
                    catch
                    {
                    }

                    try
                    {
                        Marshal.ReleaseComObject(xlApp);
                    }
                    catch
                    {
                    }

                    xlApp = null;
                }


                // =====================================================
                // 15. 남아 있는 COM RCW 정리
                // =====================================================
                GC.Collect();
                GC.WaitForPendingFinalizers();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        float GetGap(Excel.Border border)
        {
            float gap = 1f;

            if ((Excel.XlBorderWeight)border.Weight == Excel.XlBorderWeight.xlMedium)
                gap = 2f;

            if ((Excel.XlBorderWeight)border.Weight == Excel.XlBorderWeight.xlThick)
                gap = 3f;

            if ((Excel.XlLineStyle)border.LineStyle == Excel.XlLineStyle.xlDouble)
                gap += 2f;

            return gap;
        }

        private float GetBorderGap(Excel.Range cell, Excel.XlBordersIndex borderIndex)
        {
            Excel.Border? border = null;

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
            finally
            {
                if (border != null)
                    Marshal.ReleaseComObject(border);
            }
        }

        public void SnapImageMergedCell(
            Excel.Worksheet ws,
            string sheetName,
            float gapLeft = 1.5f,
            float gapTop = 1.5f,
            float gapRight = 0f,
            float gapBottom = 0.5f)
        {
            Excel.Application? xlApp = null;
            Excel.Workbook? wb = null;
            Excel.Worksheet? targetWs = null;
            Excel.Shapes? shapes = null;

            bool openedHere = false;

            try
            {
                // =====================================================
                // 1. Worksheet 준비
                // =====================================================
                if (ws == null)
                {
                    if (string.IsNullOrEmpty(xlsFilePath) ||
                        !File.Exists(xlsFilePath))
                    {
                        //MessageBox.Show(
                        //    "엑셀 파일 경로가 없습니다. xlsFilePath를 설정하세요.",
                        //    "오류",
                        //    MessageBoxButtons.OK,
                        //    MessageBoxIcon.Warning);

                        return;
                    }

                    xlApp = new Excel.Application
                    {
                        Visible = false,
                        DisplayAlerts = false
                    };

                    wb = xlApp.Workbooks.Open(
                        xlsFilePath,
                        ReadOnly: false);

                    targetWs = GetWorksheetByName(
                        wb,
                        sheetName);

                    if (targetWs == null)
                    {
                        //MessageBox.Show(
                        //    $"'{sheetName}' 시트를 찾을 수 없습니다.",
                        //    "오류",
                        //    MessageBoxButtons.OK,
                        //    MessageBoxIcon.Warning);

                        return;
                    }

                    openedHere = true;
                }
                else
                {
                    targetWs = ws;
                }


                // =====================================================
                // 2. Shapes 가져오기
                // =====================================================
                shapes = targetWs.Shapes;

                int shapeCount = shapes.Count;


                // =====================================================
                // 3. 모든 도형 처리
                // =====================================================
                for (int i = 1; i <= shapeCount; i++)
                {
                    Excel.Shape? shape = null;
                    Excel.Range? topLeft = null;
                    Excel.Range? bottomRight = null;
                    Excel.Range? area = null;

                    try
                    {
                        // -------------------------------------------------
                        // Shape
                        // -------------------------------------------------
                        shape = shapes.Item(i);

                        // -------------------------------------------------
                        // 왼쪽 위 셀
                        // -------------------------------------------------
                        topLeft = shape.TopLeftCell;


                        // =================================================
                        // 4. 병합셀 확인
                        // =================================================
                        bool isMerged = false;

                        try
                        {
                            object mergeCells = topLeft.MergeCells;

                            if (mergeCells is bool b)
                                isMerged = b;
                        }
                        catch
                        {
                            isMerged = false;
                        }


                        // =================================================
                        // 5. 실제 기준 영역 결정
                        // =================================================
                        if (isMerged)
                        {
                            // 병합셀인 경우
                            // TopLeftCell의 MergeArea 전체를 기준으로 사용
                            area = topLeft.MergeArea;
                        }
                        else
                        {
                            // 일반 셀인 경우
                            bottomRight = shape.BottomRightCell;

                            area = targetWs.Range[
                                topLeft,
                                bottomRight];
                        }


                        // =================================================
                        // 6. 회전 확인
                        // =================================================
                        double rot = 0;

                        try
                        {
                            rot = shape.Rotation % 360;

                            if (rot < 0)
                                rot += 360;
                        }
                        catch
                        {
                            rot = 0;
                        }

                        bool rotated90 =
                            Math.Abs(rot - 90) < 1.0 ||
                            Math.Abs(rot - 270) < 1.0;

                        // =================================================
                        // 테두리 보정값 계산
                        // =================================================
                        float borderLeftGap =
                            GetBorderGap(
                                area,
                                Excel.XlBordersIndex.xlEdgeLeft);

                        float borderRightGap =
                            GetBorderGap(
                                area,
                                Excel.XlBordersIndex.xlEdgeRight);

                        float borderTopGap =
                            GetBorderGap(
                                area,
                                Excel.XlBordersIndex.xlEdgeTop);

                        float borderBottomGap =
                            GetBorderGap(
                                area,
                                Excel.XlBordersIndex.xlEdgeBottom);


                        // =================================================
                        // 7. 사진 위치 및 크기 조정
                        // =================================================
                        if (rotated90)
                        {
                            // -------------------------------------------------
                            // 90 / 270도 회전된 이미지
                            // Width / Height 교체
                            // -------------------------------------------------
                            float targetW =
                                (float)area.Height;

                            float targetH =
                                (float)area.Width;


                            float newWidth =
                                targetW -
                                gapLeft -
                                gapRight;

                            float newHeight =
                                targetH -
                                gapTop -
                                gapBottom;

                            shape.Width = newWidth
                                ;

                            // 음수 방지
                            if (newWidth < 0)
                                newWidth = 0;

                            if (newHeight < 0)
                                newHeight = 0;


                            shape.Width = newWidth;
                            shape.Height = newHeight;


                            shape.Left =
                                (float)area.Left +
                                ((float)area.Width - targetW) / 2f +
                                gapLeft;

                            shape.Top =
                                (float)area.Top +
                                ((float)area.Height - targetH) / 2f +
                                gapTop;
                        }
                        else
                        {
                            // -------------------------------------------------
                            // 일반 이미지
                            // -------------------------------------------------
                            float newWidth =
                                (float)area.Width -
                                gapLeft -
                                gapRight;

                            float newHeight =
                                (float)area.Height -
                                gapTop -
                                gapBottom;


                            // 음수 방지
                            if (newWidth < 0)
                                newWidth = 0;

                            if (newHeight < 0)
                                newHeight = 0;


                            shape.Left =
                                (float)area.Left +
                                gapLeft;

                            shape.Top =
                                (float)area.Top +
                                gapTop;

                            shape.Width =
                                newWidth;

                            shape.Height =
                                newHeight;
                        }


                        // =================================================
                        // 8. 비율 고정 해제
                        // =================================================
                        shape.LockAspectRatio =
                            Microsoft.Office.Core.MsoTriState.msoFalse;


                        // =================================================
                        // 9. 셀 크기/위치 변경 시 이미지도 같이 이동
                        // =================================================
                        shape.Placement =
                            Excel.XlPlacement.xlMoveAndSize;
                    }
                    catch (Exception ex)
                    {
                        AddLog(
                            $"이미지 스냅 처리 중 오류 발생 ({i}): {ex.Message}");
                    }
                    finally
                    {
                        // =================================================
                        // COM 객체 해제
                        // =================================================

                        if (area != null)
                        {
                            Marshal.ReleaseComObject(area);
                            area = null;
                        }

                        if (bottomRight != null)
                        {
                            Marshal.ReleaseComObject(bottomRight);
                            bottomRight = null;
                        }

                        if (topLeft != null)
                        {
                            Marshal.ReleaseComObject(topLeft);
                            topLeft = null;
                        }

                        if (shape != null)
                        {
                            Marshal.ReleaseComObject(shape);
                            shape = null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AddLog(
                    $"이미지 스냅 처리 중 오류 발생: {ex.Message}");
            }
            finally
            {
                // =====================================================
                // Shapes 해제
                // =====================================================
                if (shapes != null)
                {
                    Marshal.ReleaseComObject(shapes);
                    shapes = null;
                }


                // =====================================================
                // 이 함수에서 직접 Excel을 열었다면 여기서 종료
                // =====================================================
                if (openedHere)
                {
                    // -------------------------------------------------
                    // Worksheet
                    // -------------------------------------------------
                    if (targetWs != null)
                    {
                        Marshal.ReleaseComObject(targetWs);
                        targetWs = null;
                    }


                    // -------------------------------------------------
                    // Workbook
                    // -------------------------------------------------
                    if (wb != null)
                    {
                        try
                        {
                            wb.Save();
                        }
                        catch
                        {
                        }

                        try
                        {
                            wb.Close(
                                SaveChanges: false);
                        }
                        catch
                        {
                        }

                        Marshal.ReleaseComObject(wb);
                        wb = null;
                    }


                    // -------------------------------------------------
                    // Excel Application
                    // -------------------------------------------------
                    if (xlApp != null)
                    {
                        try
                        {
                            xlApp.Quit();
                        }
                        catch
                        {
                        }

                        Marshal.ReleaseComObject(xlApp);
                        xlApp = null;
                    }


                    // -------------------------------------------------
                    // COM RCW 최종 정리
                    // -------------------------------------------------
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        #region [엑셀 시트에서 그림 삭제]
        public void RemovePictures(Excel.Worksheet? ws)
        {
            Excel.Shapes? shapes = null;

            try
            {
                shapes = ws?.Shapes;
                if (shapes == null)
                    return;

                for (int i = shapes.Count; i >= 1; i--)
                {
                    Excel.Shape? shape = null;

                    try
                    {
                        shape = shapes.Item(i);

                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture)
                        {
                            shape.Delete();
                        }
                    }
                    finally
                    {
                        if (shape != null)
                        {
                            Marshal.ReleaseComObject(shape);
                            shape = null;
                        }
                    }
                }
            }
            finally
            {
                if (shapes != null)
                {
                    Marshal.ReleaseComObject(shapes);
                    shapes = null;
                }
            }
        }

        public void RemovePictures(Excel.Worksheet ws, int? keepPictureIndex = null)
        {
            var pictureNames = new List<string>();
            Excel.Shapes shapes = ws.Shapes;

            if (keepPictureIndex == null) return;
            if (keepPictureIndex < 2) return;

            // 이미지 이름 수집
            for (int i = 1; i <= ws.Shapes.Count; i++)
            {
                Excel.Shape? shape = null;

                try
                {
                    shape = shapes.Item(i);
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                        shape.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture)
                    {
                        pictureNames.Add(shape.Name);
                    }
                }
                finally
                {
                    if (shape != null) Marshal.ReleaseComObject(shape);
                }
            }

            string? keepName = null;

            if (keepPictureIndex.HasValue &&
                keepPictureIndex.Value >= 1 &&
                keepPictureIndex.Value <= pictureNames.Count)
            {
                keepName = pictureNames[keepPictureIndex.Value - 1];
            }

            // 삭제
            for (int i = ws.Shapes.Count; i >= 1; i--)
            {
                Excel.Shape shape = ws.Shapes.Item(i);

                try
                {
                    if ((shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                         shape.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture) &&
                        shape.Name != keepName)
                    {
                        shape.Delete();
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(shape);
                }
            }
        }

        public void RemovePicturesInRange(
            Excel.Worksheet ws,
            Excel.Range targetRange,
            bool removeInside = true)
        {
            double left = (double)targetRange.Left;
            double top = (double)targetRange.Top;
            double right = left + (double)targetRange.Width;
            double bottom = top + (double)targetRange.Height;

            for (int i = ws.Shapes.Count; i >= 1; i--)
            {
                Excel.Shape shape = ws.Shapes.Item(i);

                if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoPicture &&
                    shape.Type != Microsoft.Office.Core.MsoShapeType.msoLinkedPicture)
                    continue;

                double sLeft = shape.Left;
                double sTop = shape.Top;
                double sRight = sLeft + shape.Width;
                double sBottom = sTop + shape.Height;

                bool overlap =
                    sLeft < right &&
                    sRight > left &&
                    sTop < bottom &&
                    sBottom > top;

                if ((removeInside && overlap) ||
                    (!removeInside && !overlap))
                {
                    shape.Delete();
                }
            }
        }
        #endregion

        #region [열화상 이미지 분기 시트 삽입]
        public void EnsureFeverPicturePages(
            Excel.Worksheet ws,
            int imageCount,
            int pageRowOffset = 56,
            int imageEndRow = 43,
            string lastColumn = "AD")
        {
            if (ws == null || imageCount <= 0)
                return;

            int PAGE_ROW_OFFSET = pageRowOffset;

            // 한 페이지의 사진 위치
            //const int IMAGE_START_ROW = 27;
            int IMAGE_END_ROW = imageEndRow;

            // 보고서 마지막 열
            string LAST_COLUMN = lastColumn;

            Excel.Range? printAreaRange = null;

            try
            {
                //----------------------------------------------------------
                // 1. 사진이 들어가야 하는 마지막 행 계산
                //----------------------------------------------------------

                // 페이지당 사진 2장
                int requiredPages =
                    (int)Math.Ceiling(imageCount / 2.0);

                // 예:
                // 1 page -> 43
                // 2 page -> 99
                // 3 page -> 155
                // 4 page -> 211
                int requiredImageLastRow =
                    IMAGE_END_ROW +
                    ((requiredPages - 1) * PAGE_ROW_OFFSET);


                //----------------------------------------------------------
                // 2. 현재 PrintArea의 마지막 행 확인
                //----------------------------------------------------------

                string printArea = ws.PageSetup.PrintArea;

                if (string.IsNullOrWhiteSpace(printArea))
                {
                    throw new Exception(
                        $"[{ws.Name}] 인쇄영역(PrintArea)이 설정되어 있지 않습니다.");
                }

                printAreaRange = ws.Range[printArea];

                int printStartRow = printAreaRange.Row;

                int currentPrintLastRow =
                    printAreaRange.Row +
                    printAreaRange.Rows.Count - 1;


                Debug.WriteLine(
                    $"[{ws.Name}] " +
                    $"현재 PrintArea 마지막 행={currentPrintLastRow}, " +
                    $"사진 필요 마지막 행={requiredImageLastRow}");


                //----------------------------------------------------------
                // 3. 이미 충분하면 아무것도 하지 않음
                //----------------------------------------------------------

                if (requiredImageLastRow <= currentPrintLastRow)
                {
                    Debug.WriteLine(
                        $"[{ws.Name}] 페이지 추가 필요 없음.");

                    return;
                }


                //----------------------------------------------------------
                // 4. 필요한 만큼 페이지 추가
                //----------------------------------------------------------

                while (requiredImageLastRow > currentPrintLastRow)
                {
                    Excel.Range? sourceRange = null;
                    Excel.Range? destRange = null;
                    Excel.Range? breakCell = null;

                    try
                    {
                        /*
                         * 현재 마지막 페이지 56행을 복사
                         *
                         * 예:
                         *
                         * PrintArea가 A1:AD168 이라면
                         *
                         * 마지막 페이지:
                         * 113 ~ 168
                         *
                         * 복사 위치:
                         * 169 ~ 224
                         */

                        int sourceStartRow =
                            currentPrintLastRow -
                            PAGE_ROW_OFFSET + 1;

                        int sourceEndRow =
                            currentPrintLastRow;

                        int destStartRow =
                            currentPrintLastRow + 1;

                        int destEndRow =
                            currentPrintLastRow +
                            PAGE_ROW_OFFSET;


                        Debug.WriteLine(
                            $"[{ws.Name}] 페이지 복사: " +
                            $"{sourceStartRow}~{sourceEndRow} " +
                            $"→ {destStartRow}~{destEndRow}");


                        //--------------------------------------------------
                        // 마지막 페이지 복사
                        //--------------------------------------------------

                        sourceRange = ws.Range[
                            $"A{sourceStartRow}",
                            $"{LAST_COLUMN}{sourceEndRow}"
                        ];

                        destRange = ws.Range[
                            $"A{destStartRow}",
                            $"{LAST_COLUMN}{destEndRow}"
                        ];

                        sourceRange.Copy(destRange);


                        //--------------------------------------------------
                        // 행 높이 복사
                        //
                        // Range.Copy만으로 행 높이가 정확히 복사되지
                        // 않는 경우가 있어서 별도로 맞춤
                        //--------------------------------------------------

                        for (int i = 0; i < PAGE_ROW_OFFSET; i++)
                        {
                            Excel.Range? srcRow = null;
                            Excel.Range? dstRow = null;

                            try
                            {
                                srcRow = (Excel.Range)ws.Rows[sourceStartRow + i];
                                dstRow = (Excel.Range)ws.Rows[destStartRow + i];

                                dstRow.RowHeight = srcRow.RowHeight;
                            }
                            finally
                            {
                                if (dstRow != null)
                                   ExcelComHelper.Release(dstRow);

                                if (srcRow != null)
                                    ExcelComHelper.Release(srcRow);
                            }
                        }


                        //--------------------------------------------------
                        // 새 페이지 시작 위치에 페이지 나누기
                        //--------------------------------------------------

                        //breakCell = (Excel.Range)ws.Cells[destStartRow, 1];

                        //bool breakExists = false;

                        //int breakCount = ws.HPageBreaks.Count;

                        //for (int i = 1; i <= breakCount; i++)
                        //{
                        //    Excel.HPageBreak? hp = null;
                        //    Excel.Range? location = null;

                        //    try
                        //    {
                        //        hp = ws.HPageBreaks[i];

                        //        location = hp.Location;

                        //        if (location.Row == destStartRow)
                        //        {
                        //            breakExists = true;
                        //            break;
                        //        }
                        //    }
                        //    finally
                        //    {
                        //        if (location != null)
                        //            Marshal.ReleaseComObject(location);

                        //        if (hp != null)
                        //            Marshal.ReleaseComObject(hp);
                        //    }
                        //}

                        //if (!breakExists)
                        //{
                        //    ws.HPageBreaks.Add(
                        //        Before: breakCell);
                        //}

                        breakCell = (Excel.Range)ws.Cells[destStartRow, 1];

                        bool breakExists = false;

                        Excel.HPageBreaks? hPageBreaks = null;

                        try
                        {
                            hPageBreaks = ws.HPageBreaks;

                            int breakCount = hPageBreaks.Count;

                            for (int i = 1; i <= breakCount; i++)
                            {
                                Excel.HPageBreak? hp = null;
                                Excel.Range? location = null;

                                try
                                {
                                    hp = hPageBreaks[i];
                                    location = hp.Location;

                                    if (location.Row == destStartRow)
                                    {
                                        breakExists = true;
                                        break;
                                    }
                                }
                                finally
                                {
                                    if (location != null)
                                        Marshal.ReleaseComObject(location);

                                    if (hp != null)
                                        Marshal.ReleaseComObject(hp);
                                }
                            }

                            // 없으면 추가
                            if (!breakExists)
                            {
                                hPageBreaks.Add(breakCell);
                            }
                        }
                        finally
                        {
                            if (hPageBreaks != null)
                                Marshal.ReleaseComObject(hPageBreaks);

                            if (breakCell != null)
                                Marshal.ReleaseComObject(breakCell);
                        }


                        //--------------------------------------------------
                        // PrintArea 확장
                        //--------------------------------------------------

                        currentPrintLastRow = destEndRow;

                        ws.PageSetup.PrintArea =
                            $"$A${printStartRow}:${LAST_COLUMN}${currentPrintLastRow}";


                        Debug.WriteLine(
                            $"[{ws.Name}] PrintArea 확장 → " +
                            $"A{printStartRow}:{LAST_COLUMN}{currentPrintLastRow}");
                    }
                    finally
                    {
                        if (breakCell != null)
                            ExcelComHelper.Release(breakCell);

                        if (destRange != null)
                            ExcelComHelper.Release(destRange);

                        if (sourceRange != null)
                            ExcelComHelper.Release(sourceRange);
                    }
                }


                //----------------------------------------------------------
                // 5. 최종 상태
                //----------------------------------------------------------

                Debug.WriteLine(
                    $"[{ws.Name}] 페이지 확장 완료. " +
                    $"최종 마지막 행={currentPrintLastRow}");
            }
            finally
            {
                if (printAreaRange != null)
                    ExcelComHelper.Release(printAreaRange);
            }
        }

        public void ProcFeverPicture(
            string sheetName,
            bool? bCheckBoxOcr,
            Excel.Application xlApp,
            Excel.Workbook wb,
            Excel.Workbooks wss,
            string? baseFolder,
            string? pictureFolder,
            bool? bCreateNewWs = false,
            string[]? valuesCells = null,
            int rowOffset = 27,
            int endOffet = 43,
            int rowPerPage = 56,
            string firstStartCol = "A",
            string secondStartCol = "R",
            string firstEndCol = "Q",
            string secondEndCol = "AC",
            string lastColumn = "AD")
        {
            // 새 통합문서 생성
            Excel.Workbook? newWb = null;
            Excel.Worksheet? sourceWs = null;
            Excel.Worksheet? ws = null;
            Excel.Range? rng = null;

            try
            {
                sourceWs = ExcelComHelper.GetWorksheetByLastName(wb, sheetName);
                //sourceWs = GetWorksheetByName(wb, sheetName);

                if (sourceWs == null)
                {
                    throw new Exception("분기 시트를 찾을 수 없습니다.");
                }

                if (bCreateNewWs == true)
                {
                    // 새 통합문서 생성
                    newWb = wss.Add();
                    try
                    {
                        // 기본 시트 삭제
                        //while (newWb.Worksheets.Count > 1)
                        //{
                        //    ((Excel.Worksheet)newWb.Worksheets[2]).Delete();
                        //}

                        // 분기 시트 복사
                        sourceWs.Copy(Before: newWb.Worksheets[1]);
                        //sourceWs.Activate();

                        // 기본 Sheet1 삭제
                        //((Excel.Worksheet)newWb.Worksheets[newWb.Worksheets.Count]).Delete();

                        // 복사된 시트
                        ws = (Excel.Worksheet)newWb.Worksheets[1];

                        string tempFile = Path.Combine(baseFolder ?? string.Empty, "images.xlsx");
                        newWb.SaveAs(tempFile);
                    }
                    catch { }
                }

                

                rng = sourceWs.Range["A1:R3"];

                RemovePicturesInRange(
                    sourceWs,
                    rng,
                    false);
                if (bCreateNewWs == true) RemovePictures(ws);
                //RemovePictures(sourceWs);


                string[] files = Directory.GetFiles(baseFolder + "\\" + pictureFolder, "*.jpg")
                .OrderBy(f =>
                {
                    return int.TryParse(Path.GetFileNameWithoutExtension(f), out int n)
                        ? n
                        : int.MaxValue;
                })
                .ToArray();

                // ★ 필요한 페이지가 부족하면 마지막 페이지 복사
                EnsureFeverPicturePages(sourceWs, files.Length, rowPerPage, endOffet, lastColumn);
                if (ws != null)
                {
                    EnsureFeverPicturePages(ws, files.Length, rowPerPage, endOffet, lastColumn);
                }

                int imageIndex = 0;

                for (int page = 0; imageIndex < files.Length; page++)
                {
                    int startRow = rowOffset + (page * rowPerPage);
                    int endRow = endOffet + (page * rowPerPage);

                    string[] fromCols = { firstStartCol, secondStartCol };
                    string[] toCols = { firstEndCol, secondEndCol };   // 실제 병합 끝 열에 맞게 수정

                    for (int i = 0; i < fromCols.Length; i++)
                    {
                        if (imageIndex >= files.Length)
                            break;

                        string fromCell = $"{fromCols[i]}{startRow}";
                        string toCell = $"{toCols[i]}{endRow}";

                        float gapRight = (i % 2 == 0) ? 2.3f : 2.9f;
                        float gapLeft = (i % 2 == 0) ? 4.2f : 3.6f;

                        //float gapBottom = 2.8f;
                        //float gapTop = 4.8f;

                        using (var inserter = new ImageInserter(sourceWs, files[imageIndex]))
                        {
                            inserter.InsertFit(
                                fromCell,
                                toCell,
                                new ImageInsertOptions
                                {
                                    KeepAspectRatio = false,
                                    //GapRight = 2.7f,
                                    //GapRight = gapRight,
                                    ////GapBottom = 2.6f,
                                    //GapBottom = gapBottom,
                                    ////GapLeft = 4.0f,
                                    //GapLeft = gapLeft,
                                    //GapTop = gapTop
                                });
                        }
                        if (bCreateNewWs == true && ws != null) { 

                            using (var inserter = new ImageInserter(ws, files[imageIndex]))
                            {
                                inserter.InsertFit(
                                    fromCell,
                                    toCell,
                                    new ImageInsertOptions
                                    {
                                        KeepAspectRatio = false,
                                        //GapRight = 2.7f,
                                        //GapRight = gapRight,
                                        //GapBottom = gapBottom,
                                        ////GapLeft = 4.0f,
                                        //GapLeft = gapLeft,
                                        //GapTop = gapTop
                                    });
                            }
                        }

                        imageIndex++;
                    }
                }

                if (bCheckBoxOcr == true)
                {

                    using (var reader = new FlirOcrReader())
                    {
                        OcrExcelMap map = new OcrExcelMap
                        {
                            ValueCells = valuesCells ?? new[]
                            {
                                            "H49",
                                            "P49",
                                            "W49",
                                            "H52",
                                            "P52",
                                            "W52"
                                        },
                            MinTemperatureCell = "AD6",
                            RowOffset = 56
                        };

                        string[] evenFiles = files
                            .Where((file, index) => index % 2 == 0)
                            .ToArray();

                        OcrDataToExcel.ProcessAll(
                            sourceWs,
                            evenFiles,
                            reader,
                            map);

                        if (bCreateNewWs == true) OcrDataToExcel.ProcessAll(
                            ws,
                            evenFiles,
                            reader,
                            map);
                    }
                }

                // =====================================================
                // 핵심 추가: newWb의 첫 번째 시트를 wb의 분기 시트 앞으로 복사
                // =====================================================
                //InsertFirstSheetFromNewWbBeforeBranch(wb, newWb, sheetName);


                // newWb의 첫번째 시트를 wb의 '분기'라는 이름을 포함한 시트 앞에 복사해서 넣음 (복사만, newWb에 남김)
                //InsertFirstSheetFromNewWbBeforeBranch(wb, newWb);

                AddLog($"분기 이미지 삽입 완료", "Info");
            }

            catch (Exception ex)
            {
                //throw new Exception("분기 이미지 삽입 실패: " + ex.Message, ex);
                AddLog("Error", $"분기 이미지 삽입 실패: {ex.Message}");
            }
            finally
            {
                try
                {
                    //wb.Save();
                    if (rng != null) Marshal.ReleaseComObject(rng);
                    if (ws != null)
                    {
                        Marshal.ReleaseComObject(ws);
                    }
                    if (sourceWs != null) Marshal.ReleaseComObject(sourceWs);
                    if (bCreateNewWs == true && newWb != null)
                    {
                        newWb.Save();
                        newWb.Close(false);
                        Marshal.ReleaseComObject(newWb);
                    }
                    //ExcelComHelper.Release(rng);
                    //ExcelComHelper.Release(ws);
                    //ExcelComHelper.Release(sourceWs);
                    //ExcelComHelper.CloseWorkbook(ref newWb, false);
                }
                catch
                {
                    AddLog("Error", $"분기 파일 정리 실패");
                }
            }
        }

        private void InsertFirstSheetFromNewWbBeforeBranch(
    Excel.Workbook targetWb,
    Excel.Workbook sourceWb,
    string targetSheetName)
        {
            Excel.Worksheet? newSheet = null;
            Excel.Worksheet? branchSheet = null;

            try
            {
                // 1. newWb의 첫 번째 시트 가져오기
                newSheet = sourceWb.Worksheets[1] as Excel.Worksheet;
                if (newSheet == null) return;

                // 2. targetWb에서 이름에 targetSheetName이 포함된 기준 시트 찾기
                branchSheet = ExcelComHelper.GetWorksheetByLastName(targetWb, targetSheetName);
                if (branchSheet == null) return;

                // 3. 기준 시트(branchSheet) 바로 앞으로 시트 복사 (Move 대신 Copy 사용 시 원본과 대상 양쪽에 모두 존재하게 됩니다)
                newSheet.Copy(Before: branchSheet);
            }
            catch (Exception ex)
            {
                AddLog("Error", $"시트 복사 삽입 실패: {ex.Message}");
            }
            finally
            {
                ExcelComHelper.Release(newSheet);
                ExcelComHelper.Release(branchSheet);
            }
        }

        public void ProcFeverPicture1(
            string sheetName, 
            bool? bCheckBoxOcr,
            Excel.Application xlApp,
            Excel.Workbook wb,
            Excel.Workbooks wss,
            string? baseFolder,
            string? pictureFolder)
        {
            Excel.Worksheet? sourceWs = null;
            Excel.Range? rng = null;
            try
            {
                sourceWs = ExcelComHelper.GetWorksheetByLastName(wb, sheetName);

                if (sourceWs == null)
                {
                    throw new Exception("분기 시트를 찾을 수 없습니다.");
                }

                wb.Activate();
                sourceWs.Activate();

                xlApp.ScreenUpdating = true;
                xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                // =====================================================
                // 기존 그림 제거
                // =====================================================
                //RemovePictures(sourceWs);
                rng = sourceWs.Range["A1:R3"];

                RemovePicturesInRange(
                    sourceWs,
                    rng,
                    false);

                // =====================================================
                // 이미지 파일
                // =====================================================
                string picturePath =
                    Path.Combine(baseFolder ?? string.Empty, pictureFolder ?? string.Empty);

                string[] files = Directory
                    .GetFiles(picturePath, "*.jpg")
                    .OrderBy(f =>
                    {
                        return int.TryParse(
                            Path.GetFileNameWithoutExtension(f),
                            out int n)
                            ? n
                            : int.MaxValue;
                    })
                    .ToArray();

                // =====================================================
                // 필요한 페이지 생성
                // =====================================================
                EnsureFeverPicturePages(sourceWs, files.Length);

                // =====================================================
                // 이미지 삽입
                // =====================================================
                int imageIndex = 0;

                for (int page = 0;
                     imageIndex < files.Length;
                     page++)
                {
                    int startRow = 27 + (page * 56);
                    int endRow = 43 + (page * 56);

                    string[] fromCols = { "A", "R" };
                    string[] toCols = { "Q", "AC" };

                    for (int i = 0;
                         i < fromCols.Length;
                         i++)
                    {
                        if (imageIndex >= files.Length)
                            break;

                        string fromCell =
                            $"{fromCols[i]}{startRow}";

                        string toCell =
                            $"{toCols[i]}{endRow}";

                        // =================================================
                        // 원본 통합문서
                        // =================================================
                        using (var inserter =
                            new ImageInserter(
                                sourceWs,
                                files[imageIndex]))
                        {
                            inserter.InsertFit(
                                fromCell,
                                toCell,
                                new ImageInsertOptions
                                {
                                    KeepAspectRatio = false
                                });
                        }

                        imageIndex++;
                    }
                }

                // =====================================================
                // OCR
                // =====================================================
                if (bCheckBoxOcr == true)
                {
                    using (var reader = new FlirOcrReader())
                    {
                        OcrExcelMap map = new OcrExcelMap
                        {
                            ValueCells = new[]
                                    {
                                "H49",
                                "P49",
                                "W49",
                                "H52",
                                "P52",
                                "W52"
                            },
                            MinTemperatureCell = "AD6",
                            RowOffset = 56
                        };

                        string[] evenFiles = files
                            .Where((file, index) => index % 2 == 0)
                            .ToArray();

                        OcrDataToExcel.ProcessAll(
                            sourceWs,
                            evenFiles,
                            reader,
                            map);
                    }
                }

                // =====================================================
                // 저장
                // =====================================================
                wb.Save();
                AddLog(
                    "분기 이미지 삽입 완료", "Info");
            }
            catch (Exception ex)
            {
                AddLog(
                    $"분기 이미지 삽입 실패: {ex.Message}");
            }
            finally
            {
                // -----------------------------------------------------
                // sourceWs
                // -----------------------------------------------------
                try
                {
                    if (rng != null) ExcelComHelper.Release(rng);
                    if (sourceWs != null) ExcelComHelper.Release(sourceWs);
                }
                catch { }
            }
        }
        #endregion


        #region 바닥글 이미지 교체
        private void SetRightFooterLogo(
            Excel.Worksheet ws,
            string logoPath)
        {
            if (!File.Exists(logoPath))
                throw new FileNotFoundException(
                    "로고 파일을 찾을 수 없습니다.", logoPath);

            Excel.PageSetup setup = ws.PageSetup;
            Excel.Graphic graphic = setup.RightFooterPicture;

            // 이미지 교체
            graphic.Filename = logoPath;

            // 정확한 크기를 먼저 지정하기 위해 비율 고정 해제
            graphic.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;

            // Excel은 point 단위
            // 1cm = 28.3464567 point
            graphic.Width = (float)(2.68 * 28.3464567);
            graphic.Height = (float)(0.53 * 28.3464567);

            // 가로 세로 비율 고정
            graphic.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

            // 오른쪽 바닥글에 그림 표시
            setup.RightFooter = "&G";
        }

        private void ReplaceRightFooterLogo(
            Excel.Worksheet ws,
            string logoPath)
        {
            Excel.PageSetup setup = ws.PageSetup;

            setup.RightFooter = "&G";

            Excel.Graphic graphic = setup.RightFooterPicture;

            // 기존 크기/배율 설정은 그대로 두고 이미지만 교체
            graphic.Filename = logoPath;
        }
        #endregion

        #region 갑지 이미지 중앙 정렬

        private void ReplaceA13FromSample(
            Excel.Worksheet targetWs,
            string sampleFilePath)
        {
            Excel.Application? app = null;
            Excel.Workbook? sampleWb = null;
            Excel.Worksheet? sampleWs = null;

            Excel.Range? targetCell = null;
            Excel.Range? sampleCell = null;

            Excel.Shapes? targetShapes = null;
            Excel.Shapes? sampleShapes = null;

            try
            {
                if (!File.Exists(sampleFilePath))
                    throw new FileNotFoundException(
                        "sample.xlsx 파일을 찾을 수 없습니다.",
                        sampleFilePath);

                // =====================================================
                // 대상 A13
                // =====================================================

                targetCell = targetWs.Range["A13"];

                // 병합셀이라면 MergeArea 전체를 대상으로 사용
                if (targetCell.MergeCells != null)
                {
                    Excel.Range mergeArea = targetCell.MergeArea;

                    ExcelComHelper.Release(targetCell);
                    targetCell = mergeArea;
                }


                // =====================================================
                // 현재 시트의 A13 영역에 걸쳐 있는 Shape 제거
                // =====================================================

                targetShapes = targetWs.Shapes;

                // 삭제하면서 Count가 변하므로 뒤에서부터 처리
                for (int i = targetShapes.Count; i >= 1; i--)
                {
                    Excel.Shape shape = null;

                    try
                    {
                        shape = targetShapes.Item(i);

                        double shapeLeft = shape.Left;
                        double shapeTop = shape.Top;
                        double shapeRight =
                            shape.Left + shape.Width;
                        double shapeBottom =
                            shape.Top + shape.Height;

                        double cellLeft = (double)targetCell.Left;
                        double cellTop = (double)targetCell.Top;
                        double cellRight =
                            (double)targetCell.Left + (double)targetCell.Width;
                        double cellBottom =
                            (double)targetCell.Top + (double)targetCell.Height;

                        // Shape와 A13 영역이 겹치는지 검사
                        bool overlaps =
                            shapeLeft < cellRight &&
                            shapeRight > cellLeft &&
                            shapeTop < cellBottom &&
                            shapeBottom > cellTop;

                        if (overlaps)
                        {
                            Debug.WriteLine(
                                $"A13 영역의 Shape 삭제: {shape.Name}");

                            shape.Delete();
                        }
                    }
                    finally
                    {
                        ExcelComHelper.Release(shape);
                    }
                }


                // =====================================================
                // sample.xlsx 열기
                // =====================================================

                app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                sampleWb = ExcelComHelper.OpenWorkbook(
                    app,
                    sampleFilePath,
                    true);       // ReadOnly

                sampleWs = ExcelComHelper.GetWorksheet(
                    sampleWb,
                    targetWs.Name);

                if (sampleWs == null)
                    throw new Exception(
                        $"sample.xlsx에서 '{targetWs.Name}' 시트를 찾을 수 없습니다.");


                // =====================================================
                // sample A13 셀 값 복사
                // =====================================================

                sampleCell = sampleWs.Range["A13"];

                targetCell.Value = sampleCell.Value;


                // =====================================================
                // sample A13에 걸쳐 있는 Shape 찾기
                // =====================================================

                sampleShapes = sampleWs.Shapes;

                for (int i = 1; i <= sampleShapes.Count; i++)
                {
                    Excel.Shape? sampleShape = null;
                    Excel.Range? compareCell = sampleCell;


                    try
                    {
                        sampleShape = sampleShapes.Item(i);

                        if ((bool)sampleCell.MergeCells)
                        {
                            compareCell = sampleCell.MergeArea;
                        }

                        double shapeLeft = (double)sampleShape.Left;
                        double shapeTop = (double)sampleShape.Top;
                        double shapeRight = shapeLeft + (double)sampleShape.Width;
                        double shapeBottom = shapeTop + (double)sampleShape.Height;

                        // Range 속성 캐스팅 (object -> double)
                        double cellLeft = (double)compareCell.Left;
                        double cellTop = (double)compareCell.Top;
                        double cellRight = cellLeft + (double)compareCell.Width;
                        double cellBottom = cellTop + (double)compareCell.Height;

                        bool overlaps =
                            shapeLeft < cellRight &&
                            shapeRight > cellLeft &&
                            shapeTop < cellBottom &&
                            shapeBottom > cellTop;

                        if (!overlaps)
                            continue;


                        // =================================================
                        // 이미지 복사
                        // =================================================

                        float left = sampleShape.Left;
                        float top = sampleShape.Top;
                        float width = sampleShape.Width;
                        float height = sampleShape.Height;

                        sampleShape.Copy();

                        // 현재 시트에 붙여넣기
                        targetWs.Paste();

                        // 방금 붙여넣은 Shape
                        Excel.Shape? newShape = null;

                        try
                        {
                            newShape =
                                targetWs.Shapes.Item(
                                    targetWs.Shapes.Count);

                            // sample과 동일한 위치/크기
                            newShape.Left = left;
                            newShape.Top = top;
                            newShape.Width = width;
                            newShape.Height = height;
                        }
                        finally
                        {
                            ExcelComHelper.Release(newShape);
                        }

                        // A13에 해당하는 이미지는 하나만 복사
                        break;
                    }
                    finally
                    {
                        ExcelComHelper.Release(sampleShape);
                    }
                }
            }
            finally
            {
                ExcelComHelper.Release(sampleShapes);
                ExcelComHelper.Release(targetShapes);

                ExcelComHelper.Release(sampleCell);
                ExcelComHelper.Release(targetCell);

                ExcelComHelper.Release(sampleWs);

                ExcelComHelper.CloseWorkbook(
                    ref sampleWb,
                    false);

                ExcelComHelper.QuitApplication(
                    ref app);

                ExcelComHelper.Cleanup();
            }
        }

        public ProcResult relocatePictures(string filePath, string sampleFilePath = null)
        {

            Excel.Application? xlApp = null;
            Excel.Workbook? wb = null;
            Excel.Worksheet? ws = null;
            Excel.PageSetup? pageSetup = null;
            Excel.Range? printRange = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? picture = null;
            Excel.Shape? roundRect = null;

            try
            {
                // =====================================================
                // Excel 열기
                // =====================================================

                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wb = ExcelComHelper.OpenWorkbook(
                    xlApp,
                    filePath,
                    false);

                ws = ExcelComHelper.GetWorksheet(
                    wb,
                    "갑지");

                if (ws == null)
                    return ProcResult.Fail("갑지 시트를 찾을 수 없습니다.");

                // =====================================================
                // A13 내용을 sample.xlsx 기준으로 교체
                // =====================================================

                //if (!string.IsNullOrEmpty(sampleFilePath))
                //{
                //    ReplaceA13FromSample(
                //        ws,
                //        sampleFilePath);
                //}


                // =====================================================
                // 제목 설정
                // =====================================================

                Match match =
                    Regex.Match(filePath, @"(\d{2})년(\d)분기");

                if (match.Success)
                {
                    string title =
                        $"{2000 + int.Parse(match.Groups[1].Value)}년 " +
                        $"{match.Groups[2].Value}분기" +
                        (filePath.Contains("연차") ? " 연차" : "");

                    ExcelComHelper.SetCellValue(
                        ws,
                        11,
                        1,
                        title);
                }


                // =====================================================
                // PageSetup
                // =====================================================

                pageSetup = ws.PageSetup;

                pageSetup.LeftMargin = 28.35;
                pageSetup.RightMargin = 28.35;
                pageSetup.BottomMargin = 28.35;
                pageSetup.TopMargin = 28.35;

                pageSetup.CenterHorizontally = true;
                pageSetup.CenterVertically = true;


                // =====================================================
                // 인쇄 영역
                // =====================================================

                string printArea = pageSetup.PrintArea;

                if (string.IsNullOrWhiteSpace(printArea))
                {
                    printRange = ws.UsedRange;
                }
                else
                {
                    printRange = ws.Range[printArea];
                }


                double pageLeft = (double)printRange.Left;
                double pageWidth = (double)printRange.Width;

                double centerX =
                    pageLeft + pageWidth / 2.0;

                double centerY =
                    (double)printRange.Top +
                    (double)printRange.Height / 2.0;


                // =====================================================
                // Shape 검색
                // =====================================================

                shapes = ws.Shapes;

                int shapeCount = shapes.Count;

                double minDistance = double.MaxValue;


                // =====================================================
                // 중앙에 가장 가까운 Picture 찾기
                // =====================================================

                for (int i = 1; i <= shapeCount; i++)
                {
                    Excel.Shape? shape = null;

                    try
                    {
                        shape = shapes.Item(i);

                        if (shape.Type ==
                            Microsoft.Office.Core.MsoShapeType.msoPicture)
                        {
                            double shapeCenterY =
                                shape.Top +
                                shape.Height / 2.0;

                            double distance =
                                Math.Abs(shapeCenterY - centerY);

                            if (distance < minDistance)
                            {
                                minDistance = distance;

                                // 이전 picture 해제
                                if (picture != null)
                                {
                                    ExcelComHelper.Release(picture);
                                    picture = null;
                                }

                                // 현재 shape를 picture로 넘김
                                picture = shape;
                                shape = null;
                            }
                        }
                    }
                    finally
                    {
                        ExcelComHelper.Release(shape);
                    }
                }


                // =====================================================
                // Picture 좌우 중앙 정렬
                // =====================================================

                if (picture != null)
                {
                    // picture.Shadow.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

                    picture.Left =
                        (float)(
                            centerX -
                            picture.Width / 2.0);
                }


                // =====================================================
                // 둥근 사각형 찾기
                // =====================================================

                for (int i = 1; i <= shapeCount; i++)
                {
                    Excel.Shape? shape = null;

                    try
                    {
                        shape = shapes.Item(i);

                        if (shape.Type ==
                                Microsoft.Office.Core.MsoShapeType.msoAutoShape &&
                            shape.AutoShapeType ==
                                Microsoft.Office.Core.MsoAutoShapeType
                                    .msoShapeRoundedRectangle)
                        {
                            // 호출자 소유로 넘김
                            roundRect = shape;
                            shape = null;

                            break;
                        }
                    }
                    finally
                    {
                        ExcelComHelper.Release(shape);
                    }
                }


                // =====================================================
                // 둥근 사각형 중앙 정렬
                // =====================================================

                if (roundRect != null)
                {
                    Debug.WriteLine(
                        $"PrintArea={printArea}");

                    Debug.WriteLine(
                        $"Print Left={printRange.Left}");

                    Debug.WriteLine(
                        $"Print Width={printRange.Width}");

                    Debug.WriteLine(
                        $"CenterX={centerX}");

                    Debug.WriteLine(
                        $"Before={roundRect.Left}");

                    roundRect.Left =
                        (float)(
                            centerX -
                            roundRect.Width / 2.0);

                    Debug.WriteLine(
                        $"After={roundRect.Left}");
                }


                // =====================================================
                // 저장
                // =====================================================

                wb.Save();

                return ProcResult.Ok("갑지 이미지 위치 조정이 완료되었습니다.");
            }
            catch (Exception ex)
            {
                return ProcResult.Fail(
                    $"갑지 이미지 위치 조정 중 오류가 발생했습니다.\r\n{ex.Message}");
            }
            finally
            {
                // =====================================================
                // COM 해제
                //
                // 생성한 역순으로 해제
                // =====================================================

                ExcelComHelper.Release(roundRect);
                roundRect = null;

                ExcelComHelper.Release(picture);
                picture = null;

                ExcelComHelper.Release(shapes);
                shapes = null;

                ExcelComHelper.Release(printRange);
                printRange = null;

                ExcelComHelper.Release(pageSetup);
                pageSetup = null;

                ExcelComHelper.Release(ws);
                ws = null;

                ExcelComHelper.CloseWorkbook(
                    ref wb,
                    false);

                ExcelComHelper.QuitApplication(
                    ref xlApp);

                ExcelComHelper.Cleanup();
            }
        }
        #endregion

        #region [페이지 번호 매기기]
        private int GetLastPrintRow(
            Excel.Worksheet ws,
            Excel.PageSetup pageSetup)
        {
            Excel.Range? range = null;
            Excel.Range? rows = null;

            try
            {
                string printArea =
                    pageSetup.PrintArea;

                if (!string.IsNullOrWhiteSpace(printArea))
                {
                    range = ws.Range[printArea];
                }
                else
                {
                    range = ws.UsedRange;
                }

                int firstRow = range.Row;

                rows = range.Rows;

                int rowCount = rows.Count;

                return firstRow + rowCount - 1;
            }
            finally
            {
                ExcelComHelper.Release(rows);
                ExcelComHelper.Release(range);
            }
        }

        private int GetHorizontalPageCount(
            Excel.Worksheet ws,
            int lastRow)
        {
            Excel.HPageBreaks? pageBreaks = null;

            try
            {
                pageBreaks = ws.HPageBreaks;

                int count = pageBreaks.Count;

                int pages = 1;
                int preRow = 0;

                for (int i = 1; i <= count; i++)
                {
                    Excel.HPageBreak pb = null;
                    Excel.Range location = null;

                    try
                    {
                        pb = pageBreaks.Item[i];

                        location = pb.Location;

                        int breakRow =
                            location.Row;

                        Excel.XlPageBreak breakType =
                            pb.Type;

                        // 인쇄영역 바로 다음에 존재하는
                        // 수동 페이지 나누기 무시
                        if ((breakType ==
                                 Excel.XlPageBreak.xlPageBreakManual &&
                             breakRow >= lastRow + 1)
                            ||
                            preRow >= breakRow)
                        {
                            continue;
                        }

                        Debug.WriteLine(
                            $"page : {pages}, row : {breakRow}");

                        pages++;

                        preRow = breakRow;
                    }
                    finally
                    {
                        ExcelComHelper.Release(location);
                        ExcelComHelper.Release(pb);
                    }
                }

                return pages;
            }
            finally
            {
                ExcelComHelper.Release(pageBreaks);
            }
        }

        private void DebugPageBreaks(
            Excel.Worksheet ws)
        {
            Excel.HPageBreaks pageBreaks = null;

            try
            {
                pageBreaks = ws.HPageBreaks;

                int count = pageBreaks.Count;

                for (int i = 1; i <= count; i++)
                {
                    Excel.HPageBreak pb = null;
                    Excel.Range location = null;

                    try
                    {
                        pb = pageBreaks.Item[i];
                        location = pb.Location;

                        Debug.WriteLine(
                            $"Break : {location.Address} / {pb.Type}");
                    }
                    finally
                    {
                        ExcelComHelper.Release(location);
                        ExcelComHelper.Release(pb);
                    }
                }
            }
            finally
            {
                ExcelComHelper.Release(pageBreaks);
            }
        }

        public ProcResult SetPageNumbers(string filePath)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Sheets sheets = null;

            try
            {
                if (string.IsNullOrEmpty(filePath) ||
                    !File.Exists(filePath))
                {
                    return ProcResult.Fail("엑셀 파일이 존재하지 않습니다.");
                }

                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wb = ExcelComHelper.OpenWorkbook(
                    xlApp,
                    filePath,
                    false);

                sheets = wb.Worksheets;

                int sheetCount = sheets.Count;
                int currentStartPage = 1;

                for (int i = 1; i <= sheetCount; i++)
                {
                    Excel.Worksheet? sh = null;
                    Excel.PageSetup? pageSetup = null;

                    try
                    {
                        sh = (Excel.Worksheet)sheets[i];

                        string name = sh.Name;

                        // 갑지는 제외
                        if (string.Equals(
                            name,
                            "갑지",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        sh.Activate();

                        pageSetup = sh.PageSetup;

                        // ============================================
                        // 마지막 인쇄 행
                        // ============================================

                        int lastRow =
                            GetLastPrintRow(
                                sh,
                                pageSetup);

                        // ============================================
                        // 페이지 수 계산
                        // ============================================

                        int pages =
                            GetHorizontalPageCount(
                                sh,
                                lastRow);

                        // ============================================
                        // 시작 페이지 설정
                        // ============================================

                        pageSetup.FirstPageNumber =
                            currentStartPage;

                        Debug.WriteLine(
                            $"{name} : 시작={currentStartPage}, " +
                            $"페이지수={pages}");

                        currentStartPage += pages;
                    }
                    catch (Exception ex)
                    {
                        string sheetName = "(알 수 없음)";

                        try
                        {
                            if (sh != null)
                                sheetName = sh.Name;
                        }
                        catch
                        {
                        }

                        Debug.WriteLine(
                            $"페이지 번호 처리 오류 " +
                            $"idx={i}, " +
                            $"name={sheetName}: " +
                            ex.Message);

                        // 한 시트가 실패해도 다음 시트 계속
                    }
                    finally
                    {
                        ExcelComHelper.Release(pageSetup);
                        pageSetup = null;

                        ExcelComHelper.Release(sh);
                        sh = null;
                    }
                }

                wb.Save();

                return ProcResult.Ok(
                    "페이지 번호 매기기가 완료되었습니다.");
            }
            catch (Exception ex)
            {
                return ProcResult.Fail(
                    $"페이지 번호 매기기 중 오류가 발생했습니다.\r\n{ex.Message}");
            }
            finally
            {
                ExcelComHelper.Release(sheets);
                sheets = null;

                ExcelComHelper.CloseWorkbook(
                    ref wb,
                    false);

                ExcelComHelper.QuitApplication(
                    ref xlApp);

                ExcelComHelper.Cleanup();
            }
        }
        public ProcResult SetPageStartNumbers(
            string filePath,
            Excel.Workbook wb = null,
            string firstSheetName = "갑지",
            string startSheetName = "제출문")
        {
            Excel.Application xlApp = null;
            Excel.Workbook workBook = null;
            Excel.Worksheet ws = null;
            Excel.PageSetup pageSetup = null;

            bool openedHere = false;
            bool createdExcelHere = false;

            const int xlAutomatic = -4105;

            try
            {
                // =====================================================
                // Workbook 준비
                // =====================================================
                if (wb != null)
                {
                    // 이미 열려 있는 Workbook 사용
                    workBook = wb;
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(filePath))
                        throw new ArgumentException(
                            "filePath가 없습니다.",
                            nameof(filePath));

                    if (!File.Exists(filePath))
                        throw new FileNotFoundException(
                            "엑셀 파일을 찾을 수 없습니다.",
                            filePath);

                    // =================================================
                    // Excel 실행
                    // =================================================
                    xlApp = new Excel.Application
                    {
                        Visible = false,
                        DisplayAlerts = false
                    };

                    createdExcelHere = true;

                    // =================================================
                    // 파일 열기
                    // =================================================
                    workBook = xlApp.Workbooks.Open(
                        filePath,
                        ReadOnly: false);

                    openedHere = true;
                }


                // =====================================================
                // 시트별 페이지 시작 번호 설정
                // =====================================================
                bool startFound = false;

                for (int i = 1;
                     i <= workBook.Worksheets.Count;
                     i++)
                {
                    try
                    {
                        ws = (Excel.Worksheet)workBook.Worksheets[i];
                        pageSetup = ws.PageSetup;

                        string sheetName = ws.Name;


                        // =================================================
                        // 갑지
                        // 페이지 번호 없음
                        // =================================================
                        if (sheetName == firstSheetName)
                        {
                            pageSetup.LeftFooter = "";
                            pageSetup.CenterFooter = "";
                            pageSetup.RightFooter = "";

                            continue;
                        }


                        // =================================================
                        // 제출문
                        // 시작 페이지 = 1
                        // =================================================
                        if (sheetName == startSheetName)
                        {
                            startFound = true;

                            pageSetup.FirstPageNumber = 1;

                            continue;
                        }


                        // =================================================
                        // 제출문 이후
                        // Excel의 "자동"
                        // =================================================
                        if (startFound)
                        {
                            pageSetup.FirstPageNumber =
                                xlAutomatic;
                        }
                    }
                    finally
                    {
                        if (pageSetup != null)
                        {
                            try
                            {
                                Marshal.ReleaseComObject(pageSetup);
                            }
                            catch { }

                            pageSetup = null;
                        }

                        if (ws != null)
                        {
                            try
                            {
                                Marshal.ReleaseComObject(ws);
                            }
                            catch { }

                            ws = null;
                        }
                    }
                }


                // =====================================================
                // 저장
                // =====================================================
                workBook.Save();
                return ProcResult.Ok(
                    "페이지 번호 매기기가 완료되었습니다.");
            }
            catch (Exception ex)
            {
                //AddLog(
                //    "Error",
                //    $"페이지 시작 번호 설정 실패: {ex.Message}");

                //throw;
                return ProcResult.Fail(
                    $"페이지 번호 매기기 중 오류가 발생했습니다.\r\n{ex.Message}");
            }
            finally
            {
                // =====================================================
                // 이 함수에서 직접 연 Workbook만 닫음
                // =====================================================
                if (openedHere && workBook != null)
                {
                    try
                    {
                        workBook.Close(
                            SaveChanges: true);
                    }
                    catch { }
                }

                // =====================================================
                // Workbook COM 해제
                // =====================================================
                if (workBook != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(workBook);
                    }
                    catch { }

                    workBook = null;
                }

                // =====================================================
                // 이 함수에서 만든 Excel만 종료
                // =====================================================
                if (createdExcelHere && xlApp != null)
                {
                    try
                    {
                        xlApp.Quit();
                    }
                    catch { }

                    try
                    {
                        Marshal.ReleaseComObject(xlApp);
                    }
                    catch { }

                    xlApp = null;
                }
            }
        }


        #endregion

        #region [log]
        private void AddLog(string msg, string part = "Error")
        {
            LogHandler(part, msg);
        }
        #endregion

        #region [서명 삽입]

        public void InsertInspectorSigns(
            string filePath,
            string signFolder,
            bool fitToCell = false)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;

            try
            {
                if (string.IsNullOrWhiteSpace(filePath) ||
                    !File.Exists(filePath))
                {
                    throw new FileNotFoundException(
                        "엑셀 파일을 찾을 수 없습니다.",
                        filePath);
                }

                if (string.IsNullOrWhiteSpace(signFolder) ||
                    !Directory.Exists(signFolder))
                {
                    throw new DirectoryNotFoundException(
                        $"서명 폴더가 없습니다: {signFolder}");
                }

                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wb = xlApp.Workbooks.Open(
                    filePath,
                    ReadOnly: false);

                for (int i = 1; i <= wb.Worksheets.Count; i++)
                {
                    Excel.Worksheet ws = null;

                    try
                    {
                        ws = (Excel.Worksheet)wb.Worksheets[i];

                        ProcessInspectorSignsInSheet(
                            ws,
                            signFolder,
                            fitToCell);
                    }
                    finally
                    {
                        ExcelComHelper.Release(ws);
                        ws = null;
                    }
                }

                wb.Save();

                AddLog(
                    "INFO",
                    "측정자 서명 삽입이 완료되었습니다.");
            }
            catch (Exception ex)
            {
                AddLog(
                    "ERROR",
                    $"측정자 서명 삽입 실패: {ex.Message}");

                throw;
            }
            finally
            {
                ExcelComHelper.CloseWorkbook(
                    ref wb,
                    false);
                ExcelComHelper.QuitApplication(
                    ref xlApp);
                ExcelComHelper.Cleanup();

                wb = null;
                xlApp = null;
            }
        }

        private void ProcessInspectorSignsInSheet(
            Excel.Worksheet ws,
            string signFolder,
            bool fitToCell)
        {
            Excel.Range usedRange = null;
            Excel.Range afterRange = null;
            Excel.Range found = null;

            try
            {
                usedRange = ws.UsedRange;

                afterRange = (Excel.Range)usedRange.Cells[1, 1];

                found = usedRange.Find(
                    What: "▣ 측정자 :",
                    After: afterRange,
                    LookAt: Excel.XlLookAt.xlPart,
                    LookIn: Excel.XlFindLookIn.xlValues,
                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Excel.XlSearchDirection.xlNext,
                    MatchCase: false
                );

                if (found == null)
                    return;

                string firstAddress =
                    found.Address[
                        RowAbsolute: true,
                        ColumnAbsolute: true,
                        ReferenceStyle:
                            Excel.XlReferenceStyle.xlA1];

                while (found != null)
                {
                    Excel.Range current = found;
                    found = null;

                    try
                    {
                        InsertInspectorSign(
                            ws,
                            current,
                            signFolder,
                            fitToCell);

                        // 다음 측정자 찾기
                        Excel.Range next = null;

                        try
                        {
                            next = usedRange.FindNext(current);

                            if (next == null)
                                break;

                            string nextAddress =
                                next.Address[
                                    RowAbsolute: true,
                                    ColumnAbsolute: true,
                                    ReferenceStyle:
                                        Excel.XlReferenceStyle.xlA1];

                            // 처음 찾은 셀로 돌아오면 종료
                            if (string.Equals(
                                nextAddress,
                                firstAddress,
                                StringComparison.OrdinalIgnoreCase))
                            {
                                break;
                            }

                            found = next;
                            next = null;
                        }
                        finally
                        {
                            if (next != null)
                            {
                                try
                                {
                                    Marshal.ReleaseComObject(next);
                                }
                                catch { }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        AddLog(
                            "ERROR",
                            $"[{ws.Name}] 측정자 서명 처리 실패: {ex.Message}");
                    }
                    finally
                    {
                        try
                        {
                            Marshal.ReleaseComObject(current);
                        }
                        catch { }

                        current = null;
                    }
                }
            }
            finally
            {
                ExcelComHelper.Release(afterRange);
                ExcelComHelper.Release(found);
                ExcelComHelper.Release(usedRange);
                afterRange = null;
                found = null;
                usedRange = null;
            }
        }

        private string FindInspectorSign(string text, string signFolder)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            int dashIndex = text.IndexOf('-');

            if (dashIndex < 0)
                return null;

            string namePart = text.Substring(dashIndex + 1).Trim();

            int commaIndex = namePart.IndexOf(',');
            string strName = namePart.Substring(0, (commaIndex > 0) ? commaIndex : namePart.Length).Trim();

            foreach (string signPath in Directory.GetFiles(signFolder, "*.png"))
            {
                string name = Path.GetFileNameWithoutExtension(signPath);

                if (strName.Contains(name))
                    return signPath;
            }

            return null;
        }

        private void InsertInspectorSign(
            Excel.Worksheet ws,
            Excel.Range found,
            string signFolder,
            bool fitToCell)
        {
            Excel.Range targetRange = null;
            Excel.Shapes shapes = null;
            Excel.Shape sign = null;

            try
            {
                string text =
                    Convert.ToString(found.Value2)?.Trim();

                if (string.IsNullOrWhiteSpace(text))
                    return;

                if (text.IndexOf('-') < 0)
                    return;

                string signPath =
                    FindInspectorSign(
                        text,
                        signFolder);

                if (string.IsNullOrWhiteSpace(signPath) ||
                    !File.Exists(signPath))
                {
                    AddLog(
                        "WARN",
                        $"[{ws.Name}] 측정자 서명을 찾을 수 없습니다: {text}");

                    return;
                }

                string firstName =
                    Path.GetFileNameWithoutExtension(signPath);

                if (string.IsNullOrWhiteSpace(firstName))
                    return;


                // =====================================================
                // 병합 셀 영역
                // =====================================================
                if (found.MergeCells != null)
                    targetRange = found.MergeArea;
                else
                    targetRange = found;


                // =====================================================
                // 기존 서명 삭제
                // =====================================================
                shapes = ws.Shapes;

                for (int i = shapes.Count; i >= 1; i--)
                {
                    Excel.Shape shape = null;

                    try
                    {
                        shape = shapes.Item(i);

                        if (!shape.Name.StartsWith(
                            "SIGN_",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        float centerX =
                            shape.Left +
                            shape.Width / 2;

                        float centerY =
                            shape.Top +
                            shape.Height / 2;

                        double left =
                            (double)targetRange.Left;

                        double top =
                            (double)targetRange.Top;

                        double right =
                            left +
                            (double)targetRange.Width;

                        double bottom =
                            top +
                            (double)targetRange.Height;

                        if (centerX >= left &&
                            centerX <= right &&
                            centerY >= top &&
                            centerY <= bottom)
                        {
                            shape.Delete();
                        }
                    }
                    finally
                    {
                        if (shape != null)
                        {
                            try
                            {
                                Marshal.ReleaseComObject(shape);
                            }
                            catch { }
                        }
                    }
                }


                // =====================================================
                // 문자열 뒤 서명 위치
                // =====================================================
                double fontSize =
                    Convert.ToDouble(found.Font.Size);

                double estimatedTextWidth =
                    text.Length *
                    fontSize *
                    0.72;

                const double gap = 20.0;

                double leftPosition =
                    (double)targetRange.Left +
                    estimatedTextWidth +
                    gap;

                double topPosition =
                    (double)targetRange.Top + 2;


                // =====================================================
                // 서명 삽입
                // =====================================================
                sign = ws.Shapes.AddPicture(
                    Filename: signPath,
                    LinkToFile:
                        Microsoft.Office.Core.MsoTriState.msoFalse,
                    SaveWithDocument:
                        Microsoft.Office.Core.MsoTriState.msoTrue,
                    Left: (float)leftPosition,
                    Top: (float)topPosition,
                    Width: -1,
                    Height: -1
                );

                sign.Name =
                    $"SIGN_{found.Row}_{found.Column}";

                sign.LockAspectRatio =
                    Microsoft.Office.Core.MsoTriState.msoTrue;


                // =====================================================
                // 셀에 맞추기
                // false → 원본 크기
                // true  → 셀 안에 맞게 축소
                // =====================================================
                if (fitToCell)
                {
                    double maxWidth =
                        (double)targetRange.Width -
                        estimatedTextWidth -
                        gap -
                        4;

                    double maxHeight =
                        (double)targetRange.Height - 4;

                    if (maxWidth > 0 &&
                        maxHeight > 0)
                    {
                        double scaleX =
                            maxWidth / sign.Width;

                        double scaleY =
                            maxHeight / sign.Height;

                        double scale =
                            Math.Min(
                                scaleX,
                                scaleY);

                        if (scale < 1.0)
                        {
                            sign.Width =
                                (float)(
                                    sign.Width *
                                    scale);
                        }
                    }
                }


                sign.Placement =
                    Excel.XlPlacement.xlMove;

                AddLog(
                    "INFO",
                    $"[{ws.Name}] 측정자 {firstName} 서명 삽입 " +
                    $"(크기: {(fitToCell ? "셀 맞춤" : "원본")})");
            }
            finally
            {
                ExcelComHelper.Release(sign);
                ExcelComHelper.Release(shapes);

                // found와 같은 객체면 Release하지 않음
                if (targetRange != null &&
                    targetRange != found)
                {
                    ExcelComHelper.Release(targetRange);
                }

                sign = null;
                shapes = null;
                targetRange = null;
            }
        }
        #endregion

        #region 자동페이지 나누기 검출
        public bool CheckAllPageBreaks(Excel.Workbook wb)
        {
            Excel.Sheets sheets = null;
            Excel.Worksheet ws = null;
            bool ret = false;

            try
            {
                sheets = wb.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    try
                    {
                        ws = (Excel.Worksheet)sheets[i];

                        ret = CheckPageBreaks(ws);
                    }
                    finally
                    {
                        if (ws != null)
                        {
                            ExcelComHelper.Release(ws);
                            ws = null;
                        }
                    }
                }
            }
            finally
            {
                ExcelComHelper.Release(sheets);
                sheets = null;
            }


            return ret;
        }

        private bool CheckPageBreaks(Excel.Worksheet ws)
        {
            bool ret = false;
            try
            {
                ws.DisplayPageBreaks = true;

                // 가로
                for (int i = 1; i <= ws.HPageBreaks.Count; i++)
                {
                    Excel.HPageBreak pb = null;

                    try
                    {
                        pb = ws.HPageBreaks[i];

                        if (pb.Type == Excel.XlPageBreak.xlPageBreakAutomatic)
                        {
                            string address = pb.Location?.Address ?? "";

                            AddLog("Warning",
                                $"[{ws.Name}] 자동 가로 페이지 나누기 → {address}");

                            ret = true;
                        }
                    }
                    finally
                    {
                        ExcelComHelper.Release(pb);
                        pb = null;
                    }
                }

                // 세로
                for (int i = 1; i <= ws.VPageBreaks.Count; i++)
                {
                    Excel.VPageBreak pb = null;

                    try
                    {
                        pb = ws.VPageBreaks[i];

                        if (pb.Type == Excel.XlPageBreak.xlPageBreakAutomatic)
                        {
                            string address = pb.Location?.Address ?? "";

                            AddLog("Warning",
                                $"[{ws.Name}] 자동 세로 페이지 나누기 → {address}");
                            ret = true;
                        }
                    }
                    finally
                    {
                        ExcelComHelper.Release(pb);
                        pb = null;
                    }
                }
            }
            catch (Exception ex)
            {
                AddLog("Error",
                    $"[{ws.Name}] 페이지 나누기 검사 오류: {ex.Message}");
            }
            return ret;
        }
        #endregion

        #region [측정장비 앞 공백을 측정자의 공백 개수와 맞추기]
        public void UpdateMachine(Excel.Workbook wb)
        {
            if (wb == null)
                return;

            Excel.Sheets sheets = null;

            try
            {
                sheets = wb.Worksheets;

                int count = sheets.Count;

                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet ws = null;
                    Excel.Range cells = null;
                    Excel.Range found = null;

                    try
                    {
                        ws = (Excel.Worksheet)sheets[i];
                        cells = ws.Cells;

                        found = cells.Find(
                            What: "▣ 측정장비",
                            LookAt: Excel.XlLookAt.xlPart,
                            LookIn: Excel.XlFindLookIn.xlValues);

                        if (found != null && found.Column == 1)
                        {
                            string text = found.Value2?.ToString() ?? "";
                            found.Value2 = " " + text.TrimStart();
                        }
                    }
                    finally
                    {
                        ExcelComHelper.Release(found);
                        ExcelComHelper.Release(cells);
                        ExcelComHelper.Release(ws);
                    }
                }
            }
            finally
            {
                ExcelComHelper.Release(sheets);
            }
        }
        #endregion

        #region [측정자 셀 업데이트]
        public void UpdateInspectorCells(
            Excel.Workbook wb,
            string inspectorName,
            IEnumerable<KeyValuePair<string, string>> sheetCellList)
        {

            if (string.IsNullOrEmpty(inspectorName))
            {
            //    MessageBox.Show(
            //        "측정자 이름이 null입니다. 측정자 이름을 입력하세요.",
            //        "오류",
            //        MessageBoxButtons.OK,
            //        MessageBoxIcon.Warning);

                return;
            }

            if (wb == null)
                return;

            var pairs = sheetCellList.ToList();

            // 절연 시트 존재 여부를 먼저 확인
            bool isJeolyeonSheet =
                ExcelComHelper.HasSheet(wb, "절연");

            foreach (var pair in pairs)
            {
                string sheetName = pair.Key;
                string cellAddress = pair.Value;

                Excel.Worksheet? ws = null;
                Excel.Range? range = null;

                try
                {
                    if (string.IsNullOrEmpty(sheetName))
                    {
                        // ActiveSheet 사용
                        ws = (Excel.Worksheet)wb.ActiveSheet;
                    }
                    else
                    {
                        // 해당 시트가 없으면 건너뜀
                        ws = GetWorksheetByName(wb, sheetName);
                    }

                    if (ws == null)
                        continue;

                    if (string.IsNullOrEmpty(cellAddress))
                        continue;

                    range = ws.Range[cellAddress];

                    // 절연 시트가 존재하고,
                    // 현재 대상이 절연 시트가 아니면 절연!A4 참조
                    if (isJeolyeonSheet && ws.Name != "절연")
                    {
                        range.Formula = "='절연'!A4";
                    }
                    else
                    {
                        range.Value2 = inspectorName;
                    }
                }
                catch (Exception ex)
                {
                    AddLog(
                        $"UpdateInspectorCells: failed to set " +
                        $"{sheetName ?? "(Active)"}!{cellAddress}: {ex.Message}");
                }
                finally
                {
                    ExcelComHelper.Release(range);
                    ExcelComHelper.Release(ws);
                }
            }
        }
        #endregion

        #region [바닥글에 로고 변경]
        public ProcResult ChangeFooterLogo(
            string filePath,
            string logoPath)
        {
            if (string.IsNullOrWhiteSpace(filePath) ||
                !File.Exists(filePath))
            {
                return ProcResult.Fail(
                    "엑셀 파일 경로가 올바르지 않습니다.");
            }

            if (string.IsNullOrWhiteSpace(logoPath) ||
                !File.Exists(logoPath))
            {
                return ProcResult.Fail(
                    $"로고 파일이 없습니다.\r\n{logoPath}");
            }

            Excel.Application app = null;
            Excel.Workbook wb = null;

            try
            {
                app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wb = app.Workbooks.Open(
                    filePath,
                    ReadOnly: false);

                Excel.Sheets? sheets = null;

                try
                {
                    sheets = wb.Worksheets;

                    int sheetCount = sheets.Count;

                    for (int i = 1; i <= sheetCount; i++)
                    {
                        Excel.Worksheet? ws = null;
                        Excel.PageSetup? pageSetup = null;
                        Excel.Graphic? graphic = null;

                        try
                        {
                            ws = (Excel.Worksheet)sheets[i];

                            if (ws.Name == "갑지")
                                continue;

                            pageSetup = ws.PageSetup;

                            graphic = pageSetup.RightFooterPicture;

                            // 바닥글 이미지 교체
                            graphic.Filename = logoPath;

                            // 정확한 크기 지정
                            graphic.LockAspectRatio =
                                Microsoft.Office.Core.MsoTriState.msoFalse;

                            // 2.68cm × 0.53cm
                            graphic.Width = 75.97f;
                            graphic.Height = 15.02f;

                            // 가로세로 비율 고정
                            graphic.LockAspectRatio =
                                Microsoft.Office.Core.MsoTriState.msoTrue;

                            // 오른쪽 바닥글에 그림 표시
                            pageSetup.RightFooter = "&G";

                            AddLog(
                                $"{ws.Name} 시트 바닥글 로고 변경",
                                "Info");
                        }
                        catch (Exception ex)
                        {
                            AddLog(
                                $"{ws?.Name ?? i.ToString()} 시트 로고 변경 실패: {ex.Message}",
                                "Error");
                        }
                        finally
                        {
                            ExcelComHelper.Release(graphic);
                            ExcelComHelper.Release(pageSetup);
                            ExcelComHelper.Release(ws);
                        }
                    }
                }
                finally
                {
                    ExcelComHelper.Release(sheets);
                }

                wb.Save();

                return ProcResult.Ok(
                    "모든 시트의 바닥글 로고를 변경했습니다.");
            }
            catch (Exception ex)
            {
                AddLog(
                    $"바닥글 로고 변경 실패: {ex.Message}",
                    "Error");

                return ProcResult.Fail(
                    $"바닥글 로고 변경 실패: {ex.Message}");
            }
            finally
            {
                ExcelComHelper.CloseWorkbook(
                    ref wb,
                    false);

                ExcelComHelper.QuitApplication(
                    ref app);

                ExcelComHelper.Cleanup();
            }
        }

        #endregion

        #region [품질 시트 처리]
        public void ProcQuantitySheet(Excel.Application xlApp, Excel.Workbook wb, string? baseFolder, string? quantityFolder,
            //string testReport, string timeGraph, string HighGraph)
            List<string> testReport, List<string> highGraph, List<string> timeGraph)
        {
            Excel.Worksheet? ws = null;
            string? pdfPath = null;

            try
            {
                AddLog("Info", $"품질 시트 처리 시작");
                ws = GetWorksheetByName(wb, "품질");

                if (ws == null)
                {
                    throw new Exception("품질 시트를 찾을 수 없습니다.");
                }

                RemovePictures(ws);

                pdfPath = Path.Combine(baseFolder, quantityFolder, "K.pdf");
                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {
                    inserter.Insert("U6", 0.8);

                    //wb.Save();
                }
                using (OcrReader ocr = new OcrReader("0123456789:.-~[]sec "))
                {
                    QualityPdfReader reader =
                        new QualityPdfReader(ocr);

                    QualityOcrResult result =
                        reader.Read(
                            pdfPath,
                            page: 0,
                            dpi: 400,
                            debugOutputDir:
                                Path.Combine(baseFolder, "OcrDebug"));

                    WriteQualityOcrResult(
                        ws,
                        result);
                }
                AddLog("Info", $"K.pdf 처리 완료");

                pdfPath = testReport
                    .Select(fileName => Path.Combine(baseFolder, quantityFolder, fileName))
                    .FirstOrDefault(File.Exists);

                if (pdfPath == null)
                {
                    throw new FileNotFoundException(
                        $"시험보고서 파일을 찾을 수 없습니다: {Path.Combine(baseFolder, quantityFolder)}");
                }

                //pdfPath = Path.Combine(baseFolder, quantityFolder, testReport);
                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {
                    inserter.InsertFit(
                        "A36",
                        "R67",
                        new ImageInsertOptions
                        {
                            CropLeft = 195,
                            CropTop = 170,
                            CropRight = 210,
                            CropBottom = 1000,
                            GapTop = 0,
                            GapBottom = 10.0f
                        });

                    //wb.Save();
                }

                AddLog("Info", $"시험보고서 삽입 완료");

                pdfPath = Path.Combine(baseFolder, quantityFolder, "E.pdf");
                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {

                    ImageInsertOptions option = new ImageInsertOptions
                    {
                        //CropLeft = 135,
                        //CropTop = 55,
                        //CropRight = 105,
                        //CropBottom = 146
                        GapTop = 0f,
                        GapBottom = 10.0f
                    };

                    // 1페이지
                    inserter.InsertFit(0, "A68", "R100", option);

                    // 2~5페이지
                    string[,] group1 =
                    {
                                { "A101", "G116" },   // 2
                                { "H101", "R116" },   // 3
                                { "A117", "G132" },   // 4
                                { "H117", "R132" }    // 5
                            };

                    // 6~9페이지
                    string[,] group2 =
                              {
                                { "A133", "G148" },   // 6
                                { "H133", "R148" },   // 7
                                { "A149", "G164" },   // 8
                                { "H149", "R164" }    // 9
                            };

                    // 10~13페이지
                    string[,] group3 =
                    {
                                { "A165", "G180" },   //10
                                { "H165", "R180" },   //11
                                { "A181", "G196" },   //12
                                { "H181", "R196" }    //13
                            };

                    // 14~17페이지
                    string[,] group4 =
                    {
                                { "A197", "G212" },   //14
                                { "H197", "R212" },   //15
                                { "A213", "G228" },   //16
                                { "H213", "R228" }    //17
                            };

                    int page = 1;

                    option.GapLeft = 2.8f;
                    option.GapTop = 2.8f;
                    option.GapRight = 2.8f;
                    option.GapBottom = 2.8f;

                    foreach (var group in new[] { group1, group2, group3, group4 })
                    {
                        for (int i = 0; i < 4 && page < 17; i++, page++)
                        {
                            inserter.InsertFit(
                                page,
                                group[i, 0],
                                group[i, 1],
                                option);
                        }
                    }

                    //wb.Save();
                }


                AddLog("Info", $"E.pdf 처리 완료");


                //pdfPath = Path.Combine(baseFolder, quantityFolder, timeGraph);
                pdfPath = timeGraph
                    .Select(fileName => Path.Combine(baseFolder, quantityFolder, fileName))
                    .FirstOrDefault(File.Exists);

                if (pdfPath == null)
                {
                    throw new FileNotFoundException(
                        $"시계열 그래프 파일을 찾을 수 없습니다: {Path.Combine(baseFolder, quantityFolder)}");
                }

                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {

                    ImageInsertOptions option = new ImageInsertOptions
                    {
                        CropLeft = 150,
                        CropTop = 145,
                        CropRight = 1330,
                        CropBottom = 150,
                        GapLeft = 5
                    };

                    // 1페이지
                    inserter.InsertFit("A230", "R261", option);
                    //inserter.Insert("A228", 1.0);

                    //wb.Save();
                }
                AddLog("Info", $"시계열그래프 처리 완료");

                pdfPath = highGraph
                    .Select(fileName => Path.Combine(baseFolder, quantityFolder, fileName))
                    .FirstOrDefault(File.Exists);

                if (pdfPath == null)
                {
                    throw new FileNotFoundException(
                        $"고조파 그래프 파일을 찾을 수 없습니다: {Path.Combine(baseFolder, quantityFolder)}");
                }

                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {

                    ImageInsertOptions option = new ImageInsertOptions
                    {
                        CropLeft = 150,
                        CropTop = 145,
                        CropRight = 1130,
                        CropBottom = 145,
                        GapLeft = 5
                    };

                    // 1페이지
                    inserter.InsertFit("A263", "R294", option);
                    //inserter.Insert("A262", 1.0);

                    // wb.Save();
                }
                AddLog("Info", $"고조파그래프 처리 완료");

                string[] files = Directory.GetFiles(baseFolder + $"\\{quantityFolder}", "*.bmp")
                    .OrderBy(f =>
                    {
                        return int.TryParse(Path.GetFileNameWithoutExtension(f), out int n)
                            ? n
                            : int.MaxValue;
                    })
                    .ToArray();

                int startRow = 295;
                int blockHeight = 8;   // 295~302 = 8행

                List<int> rows = new List<int> { 0, 2, 1, 7, 3, 5, 4, 6 };

                for (int i = 0; i < files.Length && i < 8; i++)
                {
                    int row = i / 2;    // 0~3
                    int col = i % 2;    // 0=좌, 1=우

                    int fromRow = startRow + row * blockHeight;
                    int toRow = fromRow + blockHeight - 1;

                    string fromCell = col == 0
                        ? $"A{fromRow}"
                        : $"H{fromRow}";

                    string toCell = col == 0
                        ? $"G{toRow}"
                        : $"R{toRow}";

                    using (var inserter = new ImageInserter(ws, files[rows[i]]))
                    {
                        inserter.InsertFit(
                            fromCell,
                            toCell,
                            new ImageInsertOptions
                            {
                                KeepAspectRatio = false
                            });
                    }
                }

                AddLog("Info", $"캡처이미지 처리 완료");

                wb.Save();
            }

            catch (Exception ex)
            {
                AddLog("Error", $"품질 이미지 삽입 실패: {ex.Message}");
            }
            finally
            {
                //try
                //{
                //    AddLog("Info", $"품질 이미지 삽입 완료");
                //    if (wb != null) Marshal.ReleaseComObject(wb);
                //    if (ws != null) Marshal.ReleaseComObject(ws);
                //}
                //catch
                //{
                //    AddLog("Error", $"분기 파일 정리 실패");
                //}

                try
                {
                    ExcelComHelper.CloseWorkbook(
                        ref wb,
                        false);
                    ExcelComHelper.Release(ws);

                    ExcelComHelper.Cleanup();
                    AddLog("Info", $"품질 이미지 삽입 완료");
                }
                catch
                {
                    AddLog("Error", $"분기 파일 정리 실패");
                }
            }
        }

        private void WriteQualityOcrResult(
            Excel.Worksheet ws,
            QualityOcrResult result)
        {
            if (ws == null)
                throw new ArgumentNullException(nameof(ws));

            if (result == null)
                throw new ArgumentNullException(nameof(result));

            Dictionary<string, string> cellMap =
                new Dictionary<string, string>
                {
                    { "measurement_period", "H5" },
                    { "record_interval",    "C6" },

                    { "max_power",          "C12" },
                    { "power_factor",       "D12" },

                    { "voltage_r",          "F12" },
                    { "voltage_s",          "F13" },
                    { "voltage_t",          "F14" },

                    { "voltage_thd_r",      "G12" },
                    { "voltage_thd_s",      "G13" },
                    { "voltage_thd_t",      "G14" },

                    { "current_r",          "J12" },
                    { "current_s",          "J13" },
                    { "current_t",          "J14" },

                    { "current_unbalance",  "M12" },

                    { "current_thd_r",      "N12" },
                    { "current_thd_s",      "N13" },
                    { "current_thd_t",      "N14" }
                };

            foreach (var pair in cellMap)
            {
                string name = pair.Key;
                string cellAddress = pair.Value;

                QualityOcrItem item = result[name];

                if (item == null)
                {
                   AddLog("Warning", $"[OCR] {name} 결과 없음");

                    continue;
                }

                if (!item.Success)
                {
                    AddLog("Warning", $"[OCR FAIL] {name} " +
                        $"Raw='{item.RawText}' " +
                        $"Conf={item.Confidence:P1}");
                }

                Excel.Range cell = null;

                try
                {
                    cell = ws.Range[cellAddress];

                    //Debug.WriteLine($"Name       = [{item.Name}]");
                    //Debug.WriteLine($"RawText    = [{item.RawText}]");
                    //Debug.WriteLine($"Text       = [{item.Text}]");
                    //Debug.WriteLine($"Numeric    = [{item.NumericValue}]");
                    //Debug.WriteLine($"Confidence = [{item.Confidence}]");
                    //Debug.WriteLine($"Success    = [{item.Success}]");

                    if (name == "measurement_period")
                    {
                        string formatted = FormatMeasurementPeriod(item.Text);

                        cell.Value2 = formatted;

                        AddLog(
                            $"[OCR OK] {name} → {cellAddress} = {formatted}");
                    }
                    else if (double.TryParse(
                        item.Text,
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out double value))
                    {
                        cell.Value2 = value;
                    }
                    else
                    {
                        cell.Value2 = item.Text;
                    }
                    

                    // 형식 오류 → 노란색
                    if (!item.FormatValid)
                    {
                        cell.Interior.Color =
                            ColorTranslator.ToOle(Color.Yellow);
                    }

                    AddLog("Info", $"[OCR OK] {name} " +
                        $"→ {cellAddress} = {item.Text} " +
                        $"({item.Confidence:P1})");
                }
                finally
                {
                    ExcelComHelper.Release(cell);
                }
            }
        }

        private string FormatMeasurementPeriod(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            // 예:
            // 2026.09.22[09:08:20] ~ 2026.09.22[11:07:40]
            //
            // 날짜와 시:분까지만 추출
            Match match = Regex.Match(
                text,
                @"(\d{4}\.\d{2}\.\d{2})\s*\[\s*(\d{2}:\d{2})(?::\d{2})?\s*\]\s*~\s*" +
                @"(\d{4}\.\d{2}\.\d{2})\s*\[\s*(\d{2}:\d{2})(?::\d{2})?\s*\]");

            if (!match.Success)
                return text;

            string startDate = match.Groups[1].Value;
            string startTime = match.Groups[2].Value;
            string endDate = match.Groups[3].Value;
            string endTime = match.Groups[4].Value;

            return $"{startDate}   {startTime}    ~    {endDate}   {endTime}";
        }
        #endregion

        #region [영코 시트 처리]
        public void ProcCoronaSheet(Excel.Application xlApp, Excel.Workbook wb, string baseFolder, string text)
        {
            string tmpFolder = Path.Combine(baseFolder, text);

            string xlsPath = Directory.GetFiles(tmpFolder, "*.xls").FirstOrDefault();

            // PD 부분방전 결과지가 있으면 PD부분방전 처리, 없으면 영코 처리
            if (xlsPath != null)
            {
                // PD부분방전
                ProcPdCoronaSheet(xlApp, wb, baseFolder, text, xlsPath);
                return;
            }

            //
            List<string> fromCols = new List<string> { "A", "R" };
            List<string> toCols = new List<string> { "Q", "AC" };
            if (ProcCry8025CoronaSheet("영코", false, xlApp, wb, baseFolder, text, 9, 10, 43, 6, fromCols, toCols))
                return;

            // 영코
            ProcVideoCoronaSheet(xlApp, wb, baseFolder, text);
        }

        private void ProcVideoCoronaSheet(
            Excel.Application xlApp,
            Excel.Workbook wb,
            string baseFolder,
            object text)
        {
            Excel.Worksheet? ws = null;
            Excel.Range? rng = null;
            Excel.PageSetup? pageSetup = null;

            try
            {
                string tmpFolder = Path.Combine(
                    baseFolder,
                    "05 영상코로나 또는 부분방전");

                string? pdfPath = Directory
                    .GetFiles(tmpFolder, "*.pdf")
                    .FirstOrDefault();

                if (pdfPath == null)
                {
                    throw new FileNotFoundException(
                        "PDF 파일을 찾을 수 없습니다.",
                        tmpFolder);
                }

                ws = GetWorksheetByName(wb, "영코");

                if (ws == null)
                {
                    throw new Exception("영코 시트를 찾을 수 없습니다.");
                }

                // =====================================================
                // 기존 이미지 삭제
                // =====================================================
                rng = ws.Range["A1:I1"];

                RemovePicturesInRange(
                    ws,
                    rng,
                    false);

                Marshal.ReleaseComObject(rng);
                rng = null;


                // =====================================================
                // PDF 이미지 삽입
                // =====================================================
                int lastImageEndRow = 24;

                using (ImageInserter inserter =
                       new ImageInserter(ws, pdfPath))
                {
                    ImageInsertOptions option =
                        new ImageInsertOptions
                        {
                            CropLeft = 175,
                            CropTop = 195,
                            CropRight = 1200,
                            CropBottom = 195
                        };

                    // 첫 번째 이미지
                    inserter.InsertFit(
                        0,
                        "A5",
                        "I24");

                    // 나머지 이미지
                    for (int i = 1;
                         i < inserter.ImageCount;
                         i++)
                    {
                        int startRow =
                            25 + (i - 1) * 25;

                        int endRow =
                            49 + (i - 1) * 25;

                        string cellFrom =
                            $"A{startRow}";

                        string cellTo =
                            $"I{endRow}";

                        inserter.InsertFit(
                            i,
                            cellFrom,
                            cellTo);

                        lastImageEndRow = endRow;
                    }

                    // 이미지가 1개인 경우도 처리
                    if (inserter.ImageCount > 0)
                    {
                        lastImageEndRow =
                            24 +
                            (inserter.ImageCount - 1) * 25;
                    }
                }


                // =====================================================
                // 페이지 설정
                // =====================================================
                pageSetup = ws.PageSetup;

                pageSetup.LeftMargin =
                    xlApp.CentimetersToPoints(2.06);

                // 마지막 이미지까지 인쇄 영역 설정
                pageSetup.PrintArea =
                    $"$A$1:$I${lastImageEndRow}";


                // 필요하다면 페이지 나누기 기준도 설정 가능
                // 페이지당 25행 구조라면 마지막 이미지까지 자동으로 포함됨.

                AddHorizontalPageBreaks(
                    ws,
                    offset: 24,
                    pageRows: 25,
                    lastRow: lastImageEndRow);


                // =====================================================
                // 저장
                // =====================================================
                wb.Save();
            }
            catch (Exception ex)
            {
                AddLog(
                    "Error",
                    $"영상코로나 이미지 삽입 실패: {ex.Message}");
            }
            finally
            {
                // Range
                ExcelComHelper.Release(rng);
                // PageSetup
                ExcelComHelper.Release(pageSetup);
                // Worksheet
                ExcelComHelper.Release(ws);
                ws = null;
                rng = null;
                pageSetup = null;
            }
        }
        private void AddHorizontalPageBreaks(
            Excel.Worksheet ws,
            int offset,
            int pageRows,
            int lastRow)
        {
            if (ws == null)
                throw new ArgumentNullException(nameof(ws));

            if (offset < 1)
                throw new ArgumentOutOfRangeException(nameof(offset));

            if (pageRows < 1)
                throw new ArgumentOutOfRangeException(nameof(pageRows));

            if (lastRow <= offset)
                return;

            try
            {
                // 기존 수동 페이지 구분선 제거
                ws.ResetAllPageBreaks();

                // 첫 페이지 다음부터 페이지 구분
                // 예:
                // offset = 24
                // pageRows = 25
                //
                // 25행
                // 50행
                // 75행
                // ...
                for (int breakRow = offset + 1;
                     breakRow <= lastRow;
                     breakRow += pageRows)
                {
                    Excel.Range breakRange = null;

                    try
                    {
                        breakRange = ws.Range[$"A{breakRow}"];

                        ws.HPageBreaks.Add(
                            breakRange);
                    }
                    finally
                    {
                        if (breakRange != null)
                        {
                            try
                            {
                                Marshal.ReleaseComObject(breakRange);
                            }
                            catch { }

                            breakRange = null;
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
        }
         
        // [Cry8025 Corona]
        public bool ProcCry8025CoronaSheet(
            string sheetName,
            bool bCheckBoxOcr,
            Excel.Application xlApp,
            Excel.Workbook wb,
            string baseFolder,
            string pictureFolder = "",
            int start = 27,
            int nRows = 16,
            int pageRows = 56,
            int nImagePerPage = 2,
            List<string> fromCols = null,
            List<string> toCols = null)
        {
            Excel.Worksheet sourceWs = null;
            if (fromCols == null)
                fromCols = new List<string> { "A", "R" };
            if (toCols == null)
                toCols = new List<string> { "Q", "AC" };

            try
            {
                sourceWs = ExcelComHelper.GetWorksheetByLastName(wb, sheetName);

                if (sourceWs == null)
                {
                    throw new Exception("영코 시트를 찾을 수 없습니다.");
                }

                // =====================================================
                // 기존 그림 제거
                // =====================================================
                RemovePictures(sourceWs);

                // =====================================================
                // 이미지 파일
                // =====================================================
                string picturePath =
                    Path.Combine(baseFolder, pictureFolder);

                string[] files = Directory
                    .GetFiles(picturePath, "*.png")
                    .OrderBy(f =>
                    {
                        return int.TryParse(
                            Path.GetFileNameWithoutExtension(f),
                            out int n)
                            ? n
                            : int.MaxValue;
                    })
                    .ToArray();

                if (files.Length == 0)
                {
                    return false;
                }

                // =====================================================
                // 필요한 페이지 생성
                // =====================================================
                // EnsureFeverPicturePages(sourceWs, files.Length);

                // =====================================================
                // 이미지 삽입
                // =====================================================
                int imageIndex = 0;

                for (int page = 0;
                     imageIndex < files.Length;)
                {
                    int startRow = start + (page * pageRows);
                    int endRow = start + nRows + (page * pageRows);

                    for (int i = 0;
                         i < nImagePerPage;
                         i++)
                    {
                        if (imageIndex >= files.Length)
                            break;

                        string fromCell =
                            $"{fromCols[i%2]}{startRow}";

                        string toCell =
                            $"{toCols[i%2]}{endRow}";

                        // =================================================
                        // 원본 통합문서
                        // =================================================
                        using (var inserter =
                            new ImageInserter(
                                sourceWs,
                                files[imageIndex]))
                        {
                            inserter.InsertFit(
                                fromCell,
                                toCell,
                                new ImageInsertOptions
                                {
                                    KeepAspectRatio = false
                                });
                        }

                        imageIndex++;
                    }
                    if (imageIndex%nImagePerPage == 0)
                    {
                        page++;
                    }
                }

                // =====================================================
                // 저장
                // =====================================================
                wb.Save();
                return true;
            }
            catch (Exception ex)
            {
                AddLog(
                    $"영코 Cry8025 이미지 삽입 실패: {ex.Message}");
                return false;
            }
            finally
            {
                AddLog(
                    "영코 Cry8025 이미지 삽입 완료", "Info");
                // -----------------------------------------------------
                // sourceWs
                // -----------------------------------------------------
                if (sourceWs != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(sourceWs);
                    }
                    catch { }

                    sourceWs = null;
                }
            }
        }

        // [PD부분방전 시트 처리]
        private Excel.Worksheet GetOrCopyWorksheet(
            Excel.Application xlApp,
            Excel.Workbook wb,
            string sheetName,
            string refPath)
        {
            Excel.Worksheet ws = null;
            Excel.Workbook refWb = null;
            Excel.Worksheet refWs = null;

            try
            {
                // 1. 현재 Workbook에서 먼저 찾기
                ws = GetWorksheetByName(wb, sheetName);

                if (ws != null)
                    return ws;

                AddLog("Info", $"{sheetName} 시트가 없어 ref.xlsx에서 가져옵니다.");

                // 2. ref.xlsx 확인
                if (!File.Exists(refPath))
                    throw new FileNotFoundException(
                        $"참조 파일을 찾을 수 없습니다: {refPath}");

                refWb = xlApp.Workbooks.Open(
                    refPath,
                    ReadOnly: true);

                refWs = GetWorksheetByName(refWb, sheetName);

                if (refWs == null)
                    throw new Exception(
                        $"ref.xlsx에도 {sheetName} 시트가 없습니다.");

                // 3. 현재 Workbook 마지막 시트 뒤에 복사
                Excel.Worksheet lastSheet = null;

                try
                {
                    lastSheet = (Excel.Worksheet)wb.Worksheets[wb.Worksheets.Count];

                    refWs.Copy(After: lastSheet);
                }
                finally
                {
                    if (lastSheet != null)
                        Marshal.ReleaseComObject(lastSheet);
                }

                // 4. 복사된 시트를 다시 가져오기
                ws = GetWorksheetByName(wb, sheetName);

                if (ws == null)
                    throw new Exception(
                        $"{sheetName} 시트 복사 후 시트를 찾지 못했습니다.");

                AddLog("Info", $"{sheetName} 시트를 ref.xlsx에서 복사했습니다.");

                return ws;
            }
            finally
            {
                ExcelComHelper.Release(refWs);
                ExcelComHelper.CloseWorkbook(ref refWb, false);
                refWs = null;
                refWb = null;
            }
        }
        private void ProcPdCoronaSheet(Excel.Application xlApp, Excel.Workbook wb, string baseFolder, string pdFolder, string xlsPath)
        {
            Excel.Worksheet ws = null;

            try
            {

                if (string.IsNullOrEmpty(xlsPath) || string.IsNullOrEmpty(xlsSamplePath))
                {
                    throw new Exception("파일을 찾을 수 없습니다.");
                }

                string tmpFolder = Path.Combine(baseFolder, pdFolder);


                ws = GetOrCopyWorksheet(
                    xlApp,
                    wb,
                    "PD부분방전",
                    xlsSamplePath);


                // PD 부분방전 결과지가 있으므로 PD 부분 방전 처리를 했다는 뜻으로 변환 과정이 실패하더라도 모두 true를 반환하도록 함

                //ws = report.GetWorksheetByName(wb, "PD부분방전");

                if (ws == null)
                {
                    throw new Exception("PD부분방전 시트를 찾을 수 없습니다.");
                }

                RemovePictures(ws, 1);


                SetPrintAreaForPDCorona(xlsPath);

                ExportExcelToPdf(xlsPath);

                // xls를 pdf로 변환한 파일을 찾음
                string pdfPath = Directory.GetFiles(tmpFolder, "*.pdf").FirstOrDefault();

                if (pdfPath == null)
                {
                    throw new Exception($"PD 부분방전 엑셀 파일을 pdf로 변환하지 못했습니다.: {tmpFolder}");
                }

                using (ImageInserter inserter = new ImageInserter(ws, pdfPath))
                {
                    ImageInsertOptions option = new ImageInsertOptions
                    {
                        CropLeft = 265,
                        CropTop = 350,
                        CropRight = 260,
                        CropBottom = 300,
                        GapRight = 5,
                        GapLeft = 3
                    };

                    int lastRow = 49;
                    int rowPerPage = 25;
                    int startRow = 25;

                    for (int i = 0; i < inserter.ImageCount; i++)
                    {
                        int index = startRow + i * 25;
                        string cellFrom = $"A{index}";
                        string cellTo = $"I{(rowPerPage - 1) + index}";
                        inserter.InsertFit(i,
                            cellFrom,
                            cellTo,
                            option
                            );
                        lastRow = (rowPerPage - 1) + index;
                    }

                    ws.PageSetup.PrintArea = $"$A$1:$I${lastRow}";

                    wb.Save();
                }
            }
            catch (Exception ex)
            {
                AddLog("Error", $"PD 부분방전 이미지 삽입 실패: {ex.Message}");
            }
            finally
            {
                try
                {
                    ExcelComHelper.CloseWorkbook(ref wb, false);
                    ExcelComHelper.Release(ws);
                    wb = null;
                    ws = null;
                    AddLog("Info", $"PD 부분방전 이미지 삽입 완료");
                }
                catch { }
            }
        }

        private void SetPrintAreaForPDCorona(string xlsFile)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Workbooks wss = null;
            Excel.Worksheet ws = null;

            try
            {
                xlApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                wss = xlApp.Workbooks;
                wb = wss.Open(xlsFile);

                // 첫 번째 시트
                ws = (Excel.Worksheet)wb.Worksheets[1];

                // 상·하·좌·우 여백 0cm
                ws.PageSetup.TopMargin = xlApp.CentimetersToPoints(0);
                ws.PageSetup.BottomMargin = xlApp.CentimetersToPoints(0);
                ws.PageSetup.LeftMargin = xlApp.CentimetersToPoints(0);
                ws.PageSetup.RightMargin = xlApp.CentimetersToPoints(0);

                ws.PageSetup.CenterHorizontally = true;
                ws.PageSetup.CenterVertically = true;

                // 실제 사용 중인 마지막 행
                Excel.Range usedRange = ws.UsedRange;

                int lastUsedRow =
                    usedRange.Row + usedRange.Rows.Count - 1;

                Marshal.ReleaseComObject(usedRange);

                // 50행 단위로 올림
                // 349 → 350
                // 300 → 300
                // 301 → 350
                int lastPrintRow =
                    ((lastUsedRow + 49) / 50) * 50;

                // 인쇄 영역 설정
                ws.PageSetup.PrintArea =
                    $"$A$1:$K${lastPrintRow}";

                // 기존 수동 페이지 나누기 삭제
                for (int i = ws.HPageBreaks.Count; i >= 1; i--)
                {
                    Excel.HPageBreak pb =
                        (Excel.HPageBreak)ws.HPageBreaks[i];

                    try
                    {
                        if (pb.Type == Excel.XlPageBreak.xlPageBreakManual)
                            pb.Delete();
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(pb);
                    }
                }

                // 50행마다 페이지 나누기
                // 350행까지라면 51, 101, 151, 201, 251, 301
                for (int row = 51; row <= lastPrintRow; row += 50)
                {
                    Excel.Range cell = null;

                    try
                    {
                        cell = (Excel.Range)ws.Cells[row, 1];
                        ws.HPageBreaks.Add(cell);
                    }
                    finally
                    {
                        if (cell != null)
                            Marshal.ReleaseComObject(cell);
                    }
                }

                wb.Save();
            }
            finally
            {
                ExcelComHelper.Release(ws);
                ExcelComHelper.Release(wss);
                ExcelComHelper.CloseWorkbook(ref wb, false);
                ExcelComHelper.QuitApplication(ref xlApp);
                ExcelComHelper.Cleanup();

                ws = null;
                wss = null;
                wb = null;
                xlApp = null;
            }
        }
        #endregion

        #region 엑셀 pdf 내보내기
        public string ExportExcelToPdf(string xlsPath)
        {
            var pdfPath = Path.ChangeExtension(xlsPath, ".pdf");

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Workbooks wss = null;


            try
            {
                app = new Excel.Application();
                wss = app.Workbooks;
                wb = wss.Open(xlsPath);

                wb.ExportAsFixedFormat(
                    Excel.XlFixedFormatType.xlTypePDF,
                    pdfPath);

                return pdfPath;
            }
            finally
            {
                ExcelComHelper.CloseWorkbook(ref wb, false);
                ExcelComHelper.Release(wss);
                ExcelComHelper.QuitApplication(ref app);
                ExcelComHelper.Cleanup();
                wb = null;
                wss = null;
                app = null;
            }
        }

        public class ExportResult
        {
            public bool Success { get; set; }
            public string OutputFile { get; set; }
            public string ErrorMessage { get; set; }
        }

        public ExportResult ExportPdfAsync(string filePath, bool bCheckBreak = false)
        {
            ExportResult returnValue = new ExportResult
            {
                Success = true
            };

            //var filePath = tbFileNameForFunction.Text?.Trim();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                returnValue = new ExportResult
                {
                    Success = false,
                    ErrorMessage = "내보낼 엑셀 파일을 먼저 선택하세요."
                };
                return returnValue;
            }

            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            string outFile = null;
            bool checkAutoBreak = false;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = xlApp.Workbooks.Open(filePath, ReadOnly: true);

                var dir = Path.GetDirectoryName(filePath) ?? Environment.CurrentDirectory;
                string baseName = Path.GetFileNameWithoutExtension(filePath);
                outFile = Path.Combine(dir, baseName + ".pdf");


                if (bCheckBreak)
                    checkAutoBreak = CheckAllPageBreaks(wb);

                if (checkAutoBreak)
                {
                    throw new Exception(
                        "자동 페이지 나누기가 감지, Log를 확인");
                }

                // Export entire workbook as a single PDF (모든 시트를 하나의 PDF로)
                try
                {
                    wb.ExportAsFixedFormat(
                        Excel.XlFixedFormatType.xlTypePDF,
                        outFile,
                        Excel.XlFixedFormatQuality.xlQualityStandard,
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        OpenAfterPublish: false);
                }
                catch (Exception ex)
                {
                    //throw new Exception("통합 PDF 내보내기 실패: " + ex.Message, ex);
                    AddLog("Error", $"통합 PDF 내보내기 실패: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                returnValue = new ExportResult
                {
                    Success = false,
                    ErrorMessage = $"PDF 내보내기 중 오류가 발생했습니다:\r\n{ex.Message}"
                };
            }
            finally
            {
                ExcelComHelper.CloseWorkbook(ref wb, false);
                ExcelComHelper.QuitApplication(ref xlApp);
                ExcelComHelper.Cleanup();
                wb = null;
                xlApp = null;

                returnValue.OutputFile = outFile;
            }
            return returnValue;
        }
        #endregion
    }
}
