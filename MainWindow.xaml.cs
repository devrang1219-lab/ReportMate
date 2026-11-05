using Microsoft.Win32;
using SoborSmartReport.Comm;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace SoborSmartReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        //
        private SoborLog _log = new SoborLog(@"D:\SoborSmartReport\log", "SmartReport");

        private Report? report = null;
        //private ReportService report = new ReportService();
        public MainWindow()
        {
            InitializeComponent();
            this.SizeChanged += MainWindow_SizeChanged;

            _log.LogHandler += (level, message) =>
            {
                AddLog(level, message);
            };
        }

        #region 공통 Dump

        private void DumpClick(string functionName)
        {
            string message =
                $"[{DateTime.Now:HH:mm:ss}] Click: {functionName}";

            if (richTextBox1 != null)
            {
                richTextBox1.AppendText(message + Environment.NewLine);
                richTextBox1.ScrollToEnd();
            }

            System.Diagnostics.Debug.WriteLine(message);
        }

        #endregion

        #region [로그 넣기]
        private void WriteLog(string log)
        {
            if (_log == null)
            {
                _log = new SoborLog();
            }

            if (string.IsNullOrEmpty(log))
                return;

            try
            {
                _log.Add(log);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"WriteLog failed: {ex.Message}");
            }
        }

        public void AddLog(string level, string message)
        {
            if (string.IsNullOrWhiteSpace(message))
                return;

            level = (level ?? "INFO").ToUpperInvariant();

            Brush levelBrush;

            switch (level)
            {
                case "ERROR":
                    levelBrush = Brushes.Red;
                    break;

                case "WARNING":
                    levelBrush = Brushes.DarkOrange;
                    break;

                case "SUCCESS":
                    levelBrush = Brushes.Green;
                    break;

                case "INFO":
                default:
                    levelBrush = Brushes.Black;
                    break;
            }

            // 로그 파일 기록은 UI 컨트롤과 무관하게 처리
            WriteLog($"[{level}] {message}");

            if (richTextBox1 == null)
                return;

            Action doAppend = () =>
            {
                try
                {
                    string timestamp =
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    var paragraph = new Paragraph
                    {
                        Margin = new Thickness(0)
                    };

                    paragraph.Inlines.Add(new Run($"[{timestamp}] ")
                    {
                        Foreground = Brushes.Gray
                    });

                    paragraph.Inlines.Add(new Run($"[{level}] ")
                    {
                        Foreground = levelBrush,
                        FontWeight = FontWeights.SemiBold
                    });

                    paragraph.Inlines.Add(new Run(message)
                    {
                        Foreground = Brushes.Black
                    });

                    richTextBox1.Document.Blocks.Add(paragraph);

                    // 마지막 로그로 스크롤
                    richTextBox1.ScrollToEnd();

                    // All 로그 탭 선택
                    //if (tapLog != null && all != null)
                    //{
                    //    tapLog.SelectedItem = all;
                    //}
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"AddLog failed: {ex.Message}");
                }
            };

            // WPF Dispatcher로 UI 스레드에서 실행
            if (richTextBox1.Dispatcher.CheckAccess())
            {
                doAppend();
            }
            else
            {
                richTextBox1.Dispatcher.BeginInvoke(
                    DispatcherPriority.Normal,
                    doAppend);
            }
        }

        #endregion

        #region 보고서 작업
        private void DropArea_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void DropArea_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void DropArea_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    tbQuantityFile.Text = files[0]; // 텍스트박스에 경로 대입
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일을 가져오는 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnQuantityFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "엑셀 파일을 선택하세요",
                Filter = "Excel 파일 (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm"
              + "|모든 파일 (*.*)|*.*",
                CheckFileExists = true,
                CheckPathExists = true,
                Multiselect = false,
                InitialDirectory = Directory.Exists(@"C:\") ? @"C:\" : ""
            };

            if (dlg.ShowDialog() == true)
            {
                tbQuantityFile.Text = dlg.FileName;
            }
            DumpClick(nameof(btnQuantityFile_Click));
        }

        private void btnGapjiPictureRelocate_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnGapjiPictureRelocate_Click));

        private void btnPageNumber_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnPageNumber_Click));

        private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (btnNavigateHide.Content.ToString() == ">>")
            {
                // 사이드바가 이미 숨겨져 있는 경우, 창 크기 변경 시 아무 작업도 하지 않음
                return;
            }

            // 예: 창 전체 너비가 800px 미만으로 줄어들면 사이드바 숨김
            if (this.ActualWidth < 1000)
            {
                SidebarColumn.Width = new GridLength(0);       // 공간 자체를 없애서 우측 화면이 넓어지게 함
                SidebarBorder.Visibility = Visibility.Collapsed; // 내용 숨김
            }
            else
            {
                SidebarColumn.Width = new GridLength(225);     // 원래 너비(250) 복원
                SidebarBorder.Visibility = Visibility.Visible;   // 내용 표시
            }
        }
        private void NavigateHide_Click(object sender, RoutedEventArgs e)
        {
            Button? button = (sender is Button) ? (Button)sender : null;

            if (button == null)
                return;

            if (button.Content.ToString() == "<<")
            {
                SidebarColumn.Width = new GridLength(50);
                button.Content = ">>";
                SidebarBorder.Visibility = Visibility.Collapsed; // 내용 숨김
            }
            else
            {
                SidebarColumn.Width = new GridLength(225);
                button.Content = "<<";
                SidebarBorder.Visibility = Visibility.Visible; // 내용 표시
            }
        }
        

        private void tbQuantityFile_TextChanged(object sender, EventArgs e)
        {
            string? filePath = tbQuantityFile.Text?.Trim();
            // string xlsSamplePath = txtBxSampleReport.Text?.Trim();
            string xlsSamplePath = @"D:\hkeng_ref\전기보고서_샘플_260911.xlsx";

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return;
            }

            Mouse.OverrideCursor = Cursors.Wait;
            //예 세정물류센터_26년3분기연차보고서_260711.xlsx

            // Report.ParseReport 호출하여 Report 인스턴스 생성
            try
            {
                this.report = Report.ParseReport(filePath, xlsSamplePath, _log);

            }
            catch (Exception ex)
            {
                AddLog("Error", $"보고서 파싱 중 오류: {ex.Message}");
            }
            finally
            {
                report?.LogHandler += (level, message) =>
                {
                    AddLog(level, message);
                };
                Mouse.OverrideCursor = null;
            }
        }

        private void btnExportForPdf_Click(object sender, RoutedEventArgs e)
        {
            btnExportForPdf.IsEnabled = false;
            Mouse.OverrideCursor = Cursors.Wait;

            string outputFile = string.Empty;
            //bool success = false;

            try
            {

                DumpClick(nameof(btnExportForPdf_Click));

                string? filePath = tbQuantityFile.Text?.Trim();

                if (string.IsNullOrWhiteSpace(filePath))
                {
                    AddLog("Error", "엑셀 파일을 먼저 선택하세요.");
                    return;
                }

                if (report == null)
                {
                    report = new Report();
                    //report.
                }

                var result = report.ExportPdfAsync(filePath, false);

                if (result == null)
                {
                    AddLog("Error", "파일 형식이 올바르지 않음");
                    return;
                }

                if (!result.Success)
                {
                    MessageBox.Show(
                        result.ErrorMessage,
                        "PDF 내보내기 실패",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    AddLog("Error",
                        $"통합 PDF 내보내기 실패: {result.ErrorMessage}");
                    return;
                }

                outputFile = result.OutputFile;
                //success = true;

                Process.Start(new ProcessStartInfo
                {
                    FileName = outputFile,
                    UseShellExecute = true
                });

                AddLog("SUCCESS",
                    $"통합 PDF 내보내기 성공: {outputFile}");
            }
            catch (Exception ex)
            {
                AddLog("Error", $"PDF 내보내기 오류: {ex.Message}");

                MessageBox.Show(
                    ex.Message,
                    "PDF 내보내기 오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
                btnExportForPdf.IsEnabled = true;
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(button1_Click));

        private void btnInsertSign_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnInsertSign_Click));

        private void btnChangeFooterLogo_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnChangeFooterLogo_Click));

        private void btnSeasonCorrect_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnSeasonCorrect_Click));

        private void btnChangeInspector_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnChangeInspector_Click));

        #endregion

        #region 정정 및 검증

        private void btnUpdateJuapList_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnUpdateJuapList_Click));

        private void btErrorPageUpdate_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btErrorPageUpdate_Click));

        private void btnConfirmJulyeon_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnConfirmJulyeon_Click));

        #endregion

        #region 이미지 정렬

        private void btnDeletePicture_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnDeletePicture_Click));

        private void btnCompressImages_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnCompressImages_Click));

        private void btnSnapImage_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnSnapImage_Click));

        #endregion

        #region 이미지 폴더 선택

        private void btnSelectThermalFolder_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnSelectThermalFolder_Click));

        private void btnSelectQuantityFolder_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnSelectQuantityFolder_Click));

        private void btnSelectCoronaFolder_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnSelectCoronaFolder_Click));

        private void btnSelectPictureFolder_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnSelectPictureFolder_Click));

        private void btnSelectEquipmentFolder_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnSelectEquipmentFolder_Click));

        private void btnSelectBatteryFolder_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnSelectBatteryFolder_Click));

        #endregion

        #region 이미지 삽입 및 OCR

        private void btnQuntityFileRun_Click(object sender, RoutedEventArgs e)
            => ProcInsetImage(nameof(btnQuntityFileRun_Click));

        private void ProcInsetImage(string functionName)
        {
            DumpClick(functionName);
            var filePath = tbQuantityFile.Text?.Trim();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("엑셀 파일을 먼저 선택하세요.", "오류", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Mouse.OverrideCursor = Cursors.Wait;
            Excel.Application? xlApp = null;
            Excel.Workbook? wb = null;
            Excel.Workbooks? workbooks = null;

            try
            {
                xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                //xlApp.ScreenUpdating = true;
                //xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                // =====================================================
                // Workbooks
                // =====================================================
                workbooks = xlApp.Workbooks;


                //wb = xlApp.Workbooks.Open(filePath, ReadOnly: false);\
                wb = workbooks.Open(filePath);


                string? baseFolder = Path.GetDirectoryName(filePath ?? string.Empty);
                //string? pdfPath = Path.Combine(baseFolder, "02 전원품질", "K.pdf");
                string? pdfPath = string.Empty;

                // 열화상
                if (checkBoxThermal.IsChecked == true)
                {
                    //ProcFeverPicture(xlApp, workbooks, wb, baseFolder, tbQuaterFolder.Text.Trim());
                    report?.ProcFeverPicture(tbQuaterSheet.Text.Trim(), checkBoxOcr.IsChecked, xlApp, wb, workbooks, baseFolder,
                        tbQuaterFolder.Text, checkBoxCopy.IsChecked);
                }

                // 품질
                if (checkBoxQuantity.IsChecked == true)
                {
                    report?.ProcQuantitySheet(xlApp, wb, baseFolder, tbQuantityFolder.Text, Constants.TEST_REPORT,
                         //comboBoxTimeGraph.Text, comboBoxHwaveGraph.Text);
                         Constants.HIGH_WAVE_GRAPH, Constants.TIME_WAVE_GRAPH);
                }


                // 코로나
                //if (checkBoxCorona.Checked)
                //{
                //    report.ProcCoronaSheet(xlApp, wb, baseFolder, textBoxCoronaFolder.Text);
                //}

                //// 점검사진
                //if (checkBoxPicture.Checked)
                //{
                //    ProcCheckPicture(xlApp, wb, baseFolder, textBoxPictureFolder.Text);
                //}

                //// 설치기기 (설비현황)
                //if (cbEquipment.Checked)
                //{
                //    ProcEquipment(xlApp, wb, baseFolder, tbEquipment.Text);
                //}

                //// 축전지 (축)
                //if (cbBattery.Checked)
                //{
                //    ProcBattery(xlApp, wb, baseFolder, tbBattery.Text);
                //}

                //// 3-2 접지저항
                //if (cbGround.Checked)
                //{
                //    ProcHighVoltage(xlApp, wb, baseFolder, tbGround.Text);
                //}

                //// 사업장현황 결선도
                //if (cbPlaceInfo.Checked)
                //{
                //    ProcLineMapImage(xlApp, wb, baseFolder, tbLineMap.Text);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이미지 삽입 실패:\r\n{ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                //return;
            }
            finally
            {
                try
                {
                    if (wb != null)
                    {
                        wb.Save();
                        wb.Close(false);
                        ExcelComHelper.Release(wb);
                        //ExcelComHelper.CloseWorkbook(ref wb, false);
                    }
                }
                catch { }

                if (workbooks != null)
                {
                    try
                    {
                        ExcelComHelper.Release(workbooks);
                    }
                    catch
                    {
                    }
                }

                try
                {
                    if (xlApp != null)
                    {
                        //xlApp.Quit();
                        //ExcelComHelper.Release(xlApp);
                        ExcelComHelper.QuitApplication(ref xlApp);
                    }
                }
                catch { }

                ExcelComHelper.Cleanup();
                Mouse.OverrideCursor = null;
            }
        }

        private void ProcFeverPicture(Excel.Application xlApp, Excel.Workbooks wss, Excel.Workbook wb, string? baseFolder, string? pictureFolder)
        {
            // 새 통합문서 생성
            Excel.Workbook? newWb = null;
            Excel.Worksheet? sourceWs = null;
            Excel.Worksheet? ws = null;
            Excel.Range? rng = null;

            string sheetName = tbQuaterSheet.Text.Trim();

            try
            {
                sourceWs = ExcelComHelper.GetWorksheetByLastName(wb, sheetName);

                if (sourceWs == null)
                {
                    throw new Exception("분기 시트를 찾을 수 없습니다.");
                }

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

                rng = sourceWs.Range["A1:R3"];

                //report?.RemovePicturesInRange(
                //    sourceWs,
                //    rng,
                //    false);
                report?.RemovePictures(ws);
                report?.RemovePictures(sourceWs);


                string[] files = Directory.GetFiles(baseFolder + "\\" + pictureFolder, "*.jpg")
                .OrderBy(f =>
                {
                    return int.TryParse(Path.GetFileNameWithoutExtension(f), out int n)
                        ? n
                        : int.MaxValue;
                })
                .ToArray();

                // ★ 필요한 페이지가 부족하면 마지막 페이지 복사
                report?.EnsureFeverPicturePages(sourceWs, files.Length);
                report?.EnsureFeverPicturePages(ws, files.Length);

                int imageIndex = 0;

                for (int page = 0; imageIndex < files.Length; page++)
                {
                    int startRow = 27 + (page * 56);
                    int endRow = 43 + (page * 56);

                    string[] fromCols = { "A", "R" };
                    string[] toCols = { "Q", "AC" };   // 실제 병합 끝 열에 맞게 수정

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
                                    ////GapBottom = 2.6f,
                                    //GapBottom = gapBottom,
                                    ////GapLeft = 4.0f,
                                    //GapLeft = gapLeft,
                                    //GapTop = gapTop
                                });
                        }

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
                                    //GapBottom = gapBottom,
                                    ////GapLeft = 4.0f,
                                    //GapLeft = gapLeft,
                                    //GapTop = gapTop
                                });
                        }

                        imageIndex++;
                    }
                }

                if (checkBoxOcr.IsChecked == true)
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

                        OcrDataToExcel.ProcessAll(
                            ws,
                            evenFiles,
                            reader,
                            map);
                    }
                }

                wb.Save();
                newWb.Save();

                // newWb의 첫번째 시트를 wb의 '분기'라는 이름을 포함한 시트 앞에 복사해서 넣음 (복사만, newWb에 남김)
                //InsertFirstSheetFromNewWbBeforeBranch(wb, newWb);
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
                    AddLog("Info", $"분기 이미지 삽입 완료");
                    //if (rng != null) Marshal.ReleaseComObject(rng);
                    //if (ws != null) Marshal.ReleaseComObject(ws);
                    //if (sourceWs != null) Marshal.ReleaseComObject(sourceWs);
                    //if (newWb != null)
                    //{
                    //    newWb.Close(false);
                    //    Marshal.ReleaseComObject(newWb);
                    //}
                    ExcelComHelper.Release(rng);
                    ExcelComHelper.Release(ws);
                    ExcelComHelper.Release(sourceWs);
                    ExcelComHelper.CloseWorkbook(ref newWb, false);
                    ExcelComHelper.Cleanup();
                }
                catch
                {
                    AddLog("Error", $"분기 파일 정리 실패");
                }
            }
        }

        #endregion

        #region 파일 다운로드

        private void btnDownload_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(btnDownload_Click));

        #endregion

        #region 사이드 메뉴

        private void NavigateReport_Click(object sender, RoutedEventArgs e)
        {
            mainTabs.SelectedIndex = 0;
            DumpClick(nameof(NavigateReport_Click));
        }

        private void NavigateCorrection_Click(object sender, RoutedEventArgs e)
        {
            mainTabs.SelectedIndex = 1;
            DumpClick(nameof(NavigateCorrection_Click));
        }

        private void NavigateAlign_Click(object sender, RoutedEventArgs e)
        {
            mainTabs.SelectedIndex = 2;
            DumpClick(nameof(NavigateAlign_Click));
        }

        private void NavigateInsert_Click(object sender, RoutedEventArgs e)
        {
            mainTabs.SelectedIndex = 3;
            DumpClick(nameof(NavigateInsert_Click));
        }

        private void NavigateDownload_Click(object sender, RoutedEventArgs e)
        {
            mainTabs.SelectedIndex = 4;
            DumpClick(nameof(NavigateDownload_Click));
        }

        private void NavigateLog_Click(object sender, RoutedEventArgs e)
        {
            DumpClick(nameof(NavigateLog_Click));
        }

        #endregion

        #region 로그 필터

        private void ShowAllLog_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(ShowAllLog_Click));

        private void ShowInfoLog_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(ShowInfoLog_Click));

        private void ShowErrorLog_Click(object sender, RoutedEventArgs e)
            => DumpClick(nameof(ShowErrorLog_Click));

        #endregion

        #region 탭 변경

        private void mainTabs_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            // 내부 탭 선택 이벤트가 버블링되는 경우까지 로그에 남기지 않도록
            // mainTabs 자체의 선택 변경만 기록합니다.
            if (ReferenceEquals(sender, mainTabs))
            {
                DumpClick(nameof(mainTabs_SelectionChanged));
            }
        }

        #endregion
    }
}