
using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace SoborSmartReport.Comm
{
    public static class ExcelComHelper
    {
        // ============================================================
        // COM Release
        // ============================================================

        /// <summary>
        /// COM 객체를 해제한다.
        ///
        /// 주의:
        /// 반환받은 COM 객체의 소유권이 호출자에게 있을 때 사용한다.
        /// 다른 코드와 공유 중인 COM 객체에는 호출하지 않는다.
        /// </summary>
        public static void Release(object? obj)
        {
            if (obj == null)
                return;

            try
            {
                if (Marshal.IsComObject(obj))
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch (InvalidComObjectException)
            {
                // 이미 해제된 COM 객체
            }
            catch (COMException)
            {
                // COM 해제 과정의 예외
            }
        }

        /// <summary>
        /// 여러 COM 객체를 순서대로 해제한다.
        /// </summary>
        public static void ReleaseAll(params object[] objects)
        {
            if (objects == null)
                return;

            for (int i = objects.Length - 1; i >= 0; i--)
            {
                Release(objects[i]);
            }
        }

        // ============================================================
        // Worksheet
        // ============================================================

        /// <summary>
        /// 시트 이름에 특정 문자열이 포함된 Worksheet를 반환한다.
        /// 반환된 Worksheet는 호출자가 Release 해야 한다.
        /// </summary>
        public static Excel.Worksheet GetWorksheet(
            Excel.Workbook wb,
            string sheetName)
        {
            if (wb == null || string.IsNullOrWhiteSpace(sheetName))
                return null;

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
                        sheet = sheets[i] as Excel.Worksheet;

                        if (sheet == null)
                            continue;

                        string name = sheet.Name ?? "";

                        if (name.IndexOf(
                                sheetName,
                                StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Excel.Worksheet result = sheet;
                            sheet = null;

                            return result;
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            Release(sheet);
                        }
                    }
                }

                return null;
            }
            finally
            {
                if (sheets != null)
                {
                    Release(sheets);
                }
            }
        }

        /// <summary>
        /// 시트 이름이 특정 문자열로 끝나는 Worksheet를 반환한다.
        /// 반환된 Worksheet는 호출자가 Release 해야 한다.
        /// </summary>
        public static Excel.Worksheet? GetWorksheetByLastName(
            Excel.Workbook wb,
            string sheetName)
        {
            if (wb == null || string.IsNullOrWhiteSpace(sheetName))
                return null;

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
                        sheet = sheets[i] as Excel.Worksheet;

                        if (sheet == null)
                            continue;

                        string name = sheet.Name ?? "";

                        if (name.EndsWith(
                                sheetName,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            // 호출 쪽에서 해제하도록
                            Excel.Worksheet result = sheet;
                            sheet = null;

                            return result;
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            Release(sheet);
                        }
                    }
                }

                return null;
            }
            finally
            {
                Release(sheets);
            }
        }

        // ============================================================
        // Common Sheet Checks
        // ============================================================

        public static bool HasYebiSheet(Excel.Workbook wb)
        {
            return HasSheet(wb, "예비");
        }

        public static bool HasCoronaSheet(Excel.Workbook wb)
        {
            return HasSheet(wb, "영코", "PD부분방전");
        }

        public static bool HasNaeryuckSheet(Excel.Workbook wb)
        {
            return HasSheet(wb, "내력");
        }

        /// <summary>
        /// 시트 이름에 지정된 검색어 중 하나라도 포함되어 있는지 확인한다.
        /// 대소문자를 구분하지 않는다.
        /// </summary>
        public static bool HasSheet(
            Excel.Workbook wb,
            params string[] sheetNames)
        {
            if (wb == null ||
                sheetNames == null ||
                sheetNames.Length == 0)
            {
                return false;
            }

            string[] keywords = Array.FindAll(
                sheetNames,
                name => !string.IsNullOrWhiteSpace(name));

            if (keywords.Length == 0)
                return false;

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
                        sheet = sheets[i] as Excel.Worksheet;

                        if (sheet == null)
                            continue;

                        string name = sheet.Name ?? "";

                        foreach (string keyword in keywords)
                        {
                            if (name.IndexOf(
                                    keyword,
                                    StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return true;
                            }
                        }
                    }
                    finally
                    {
                        Release(sheet);
                    }
                }

                return false;
            }
            finally
            {
                Release(sheets);
            }
        }

        // ============================================================
        // Cell
        // ============================================================

        /// <summary>
        /// 셀의 표시 문자열을 반환한다.
        /// Excel에서 보이는 형식을 기준으로 읽는다.
        /// </summary>
        public static string GetCellText(
            Excel.Worksheet ws,
            int row,
            int col)
        {
            if (ws == null || row < 1 || col < 1)
                return "";

            Excel.Range? cells = null;
            Excel.Range? cell = null;

            try
            {
                cells = ws.Cells;
                cell = (Excel.Range)cells[row, col];

                return Convert.ToString(
                    cell.Text,
                    CultureInfo.CurrentCulture)?.Trim() ?? "";
            }
            finally
            {
                if (cell != null)
                {
                    Release(cell);
                }

                if (cells != null)
                {
                    Release(cells);
                }
            }
        }

        /// <summary>
        /// 셀의 실제 값을 반환한다.
        /// 수식 셀은 Value2 기준 결과를 반환한다.
        /// </summary>
        public static object GetCellValue(
            Excel.Worksheet ws,
            int row,
            int col)
        {
            if (ws == null || row < 1 || col < 1)
                return null;

            Excel.Range cells = null;
            Excel.Range cell = null;

            try
            {
                cells = ws.Cells;
                cell = (Excel.Range)cells[row, col];

                return cell.Value2;
            }
            finally
            {
                ReleaseAll(cell, cells);
            }
        }

        public static void SetCellValue(
            Excel.Worksheet ws,
            int row,
            int col,
            object value)
        {
            if (ws == null || row < 1 || col < 1)
                return;

            Excel.Range? cells = null;
            Excel.Range? cell = null;

            try
            {
                cells = ws.Cells;
                cell = (Excel.Range)cells[row, col];

                cell.Value2 = value;
            }
            finally
            {
                if (cell != null)
                {
                    Release(cell);
                }

                if (cells != null)
                {
                    Release(cells);
                }
            }
        }

        // ============================================================
        // Range
        // ============================================================

        /// <summary>
        /// 주소로 Range를 반환한다.
        /// 반환된 Range는 호출자가 Release 해야 한다.
        /// </summary>
        public static Excel.Range GetRange(
            Excel.Worksheet ws,
            string address)
        {
            if (ws == null || string.IsNullOrWhiteSpace(address))
                return null;

            Excel.Range range = null;

            try
            {
                range = ws.Range[address];

                Excel.Range result = range;
                range = null;

                return result;
            }
            finally
            {
                Release(range);
            }
        }

        /// <summary>
        /// 시작 셀부터 끝 셀까지의 Range를 반환한다.
        /// 반환된 Range는 호출자가 Release 해야 한다.
        /// </summary>
        public static Excel.Range GetRange(
            Excel.Worksheet ws,
            string from,
            string to)
        {
            if (ws == null ||
                string.IsNullOrWhiteSpace(from) ||
                string.IsNullOrWhiteSpace(to))
            {
                return null;
            }

            Excel.Range? range = null;

            try
            {
                range = ws.Range[from, to];

                Excel.Range result = range;
                range = null;

                return result;
            }
            finally
            {
                if (range != null)
                {
                    Release(range);
                }
            }
        }

        public static string GetRangeText(
            Excel.Worksheet ws,
            string address)
        {
            if (ws == null || string.IsNullOrWhiteSpace(address))
                return "";

            Excel.Range range = null;

            try
            {
                range = ws.Range[address];

                return Convert.ToString(
                    range.Text,
                    CultureInfo.CurrentCulture)?.Trim() ?? "";
            }
            finally
            {
                Release(range);
            }
        }

        public static object GetRangeValue(
            Excel.Worksheet ws,
            string address)
        {
            if (ws == null || string.IsNullOrWhiteSpace(address))
                return null;

            Excel.Range? range = null;

            try
            {
                range = ws.Range[address];

                return range.Value2;
            }
            finally
            {
                if (range != null)
                {
                    Release(range);
                }
            }
        }

        public static void SetRangeValue(
            Excel.Worksheet ws,
            string address,
            object value)
        {
            if (ws == null || string.IsNullOrWhiteSpace(address))
                return;

            Excel.Range? range = null;

            try
            {
                range = ws.Range[address];

                range.Value2 = value;
            }
            finally
            {
                Release(range);
            }
        }

        // ============================================================
        // Workbook Open
        // ============================================================

        /// <summary>
        /// Workbook을 연다.
        /// 반환된 Workbook은 호출자가 CloseWorkbook 또는
        /// Close + Release를 수행해야 한다.
        ///
        /// Excel.Application은 호출자가 소유하며,
        /// 이 메서드에서 종료하지 않는다.
        /// </summary>
        public static Excel.Workbook OpenWorkbook(
            Excel.Application app,
            string filePath,
            bool readOnly = false)
        {
            if (app == null)
                throw new ArgumentNullException(nameof(app));

            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException(
                    "파일 경로가 비어 있습니다.",
                    nameof(filePath));

            Excel.Workbooks? workbooks = null;
            Excel.Workbook? wb = null;

            try
            {
                workbooks = app.Workbooks;

                wb = workbooks.Open(
                    Filename: filePath,
                    UpdateLinks: 0,
                    ReadOnly: readOnly,
                    IgnoreReadOnlyRecommended: true,
                    AddToMru: false);

                Excel.Workbook result = wb;
                wb = null;

                return result;
            }
            finally
            {
                // 성공 시 반환된 Workbook은 해제하지 않는다.
                Release(wb);
                Release(workbooks);
            }
        }

        // ============================================================
        // Workbook Close
        // ============================================================

        /// <summary>
        /// Workbook을 닫고 호출자 참조를 null로 만든다.
        /// saveChanges=true이면 변경 사항 저장을 요청한다.
        /// </summary>
        public static void CloseWorkbook(
            ref Excel.Workbook? wb,
            bool saveChanges = false)
        {
            if (wb == null)
                return;

            Excel.Workbook target = wb;
            wb = null;

            try
            {
                target.Close(SaveChanges: saveChanges);
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Excel Workbook.Close 오류: {ex.Message}");
            }
            finally
            {
                Release(target);
            }
        }

        // ============================================================
        // Application Quit
        // ============================================================

        /// <summary>
        /// Excel.Application을 종료하고 참조를 null로 만든다.
        ///
        /// 주의:
        /// 다른 Workbook이나 코드에서 공유하는 Application에는
        /// 호출하지 않는다.
        /// </summary>
        public static void QuitApplication(
            ref Excel.Application app)
        {
            if (app == null)
                return;

            Excel.Application target = app;
            app = null;

            try
            {
                target.Quit();
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Excel Application.Quit 오류: {ex.Message}");
            }
            finally
            {
                Release(target);
            }
        }

        // ============================================================
        // GC
        // ============================================================

        /// <summary>
        /// 필요할 때만 명시적으로 호출한다.
        /// 일반적인 셀 읽기/쓰기마다 호출하지 않는다.
        /// </summary>
        public static void Cleanup()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
