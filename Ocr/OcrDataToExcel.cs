using OpenCvSharp;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace SoborSmartReport
{
    public static class OcrDataToExcel
    {
        /// <summary>
        /// 이미지 한 장 처리
        /// </summary>
        public static void Process(
            Excel.Worksheet ws,
            FlirResult result,
            OcrExcelMap map,
            int rowOffset)
        {
            int index = 0;

            foreach (var item in result.Items.OrderBy(x => x.Key))
            {
                if (index >= map.ValueCells.Length)
                    break;

                string cell = AddRow(map.ValueCells[index], rowOffset);

                ws.Range[cell].Value = item.Value;

                index++;
            }
        }

        /// <summary>
        /// 이미지 여러 장 처리
        /// </summary>
        public static void ProcessAll(
            Excel.Worksheet ws,
            string[] imageFiles,
            FlirOcrReader reader,
            OcrExcelMap map)
        {
            double minTemperature = double.MaxValue;

            for (int i = 0; i < imageFiles.Length; i++)
            {
                FlirResult result = reader.Read(imageFiles[i]);

                Debug.WriteLine($"Image : {Path.GetFileName(imageFiles[i])}");
                Debug.WriteLine($"Item Count : {result.Items.Count}");
                Debug.WriteLine($"Scale Min : {result.ScaleMinTemperature}");
                Debug.WriteLine($"Left OCR : {result.LeftRawText}");
                Debug.WriteLine($"Right OCR : {result.RightRawText}");

                //foreach (var item in result.Items)
                //{
                //    Debug.WriteLine($"{item.Key} = {item.Value}");
                //}

                Process(ws, result, map, i * map.RowOffset);

                if (!double.IsNaN(result.ScaleMinTemperature))
                {
                    string cell = AddRow(map.MinTemperatureCell, i * map.RowOffset);

                    // 환경온도(컬러바 하단)의 최소값을 넣음
                    minTemperature = Math.Min(minTemperature, result.ScaleMinTemperature);

                    // cell 의 바로 위 행의 값이 비어있으면 continue
                    // 바로 위 셀 주소
                    string upperCell = AddRow(map.MinTemperatureCell, i * map.RowOffset - 1);

                    // 바로 위 셀이 있으면 입력
                    if (!string.IsNullOrWhiteSpace(Convert.ToString(ws.Range[upperCell].Value)))
                        ws.Range[cell].Value = result.ScaleMinTemperature;
                }
            }

            if (minTemperature != double.MaxValue)
            {
                ws.Range[map.MinTemperatureCell].Value = Math.Ceiling(minTemperature);
                //'옥내(사용온도 2℃)
                ws.Range["T6"].Value = @"'옥내(사용온도 " + Math.Ceiling(minTemperature) + "℃)";
            }
        }

        private static string AddRow(string cell, int offset)
        {
            int pos = 0;

            while (pos < cell.Length && char.IsLetter(cell[pos]))
                pos++;

            string col = cell.Substring(0, pos);
            int row = int.Parse(cell.Substring(pos));

            return col + (row + offset);
        }
    }

    public class OcrExcelMap
    {
        // 한 이미지당 최대 6개의 측정값을 넣을 셀
        public string[] ValueCells { get; set; }

        // 최저 환경온도(컬러바 하단)의 최소값을 넣을 셀
        public string MinTemperatureCell { get; set; }

        // 다음 이미지가 시작되는 행 간격
        public int RowOffset { get; set; }
    }

}
