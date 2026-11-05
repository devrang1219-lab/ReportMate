
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Threading;

namespace SoborSmartReport.Comm
{
    public class SoborLog : IDisposable
    {
        // 설정
        public string PathBase = @"C:\Soborsoft\log\";
        public int Interval = 2000;
        public int Period = 365;
        public bool DoSave = true;
        public bool DoDeleteOldLog = true;
        public bool UseLineHeader = true;

        public string Tag = "";

        // 로그 큐
        private readonly ConcurrentQueue<string> logList =
            new ConcurrentQueue<string>();

        // WPF 로그 출력 컨트롤
        private ListBox allLog;
        private ListBox errorLog;
        private ListBox warningLog;
        private ListBox infoLog;

        // 비동기 루프 관리
        private CancellationTokenSource _cancellationTokenSource;
        private Task _loopTask = Task.CompletedTask;
        private readonly object _loopLock = new object();
        private bool _disposed;

        // 마지막 로그 정리 날짜
        private string _lastCleanupDate = "";

        private const int MaxLogLines = 200;

        public Action<string, string> LogHandler { get; set; }


        private string LogPath
        {
            get
            {
                return Path.Combine(
                    PathBase,
                    DateTime.Now.ToString("yyyy"),
                    DateTime.Now.ToString("MM"));
            }
        }

        private string LogFile
        {
            get
            {
                string fileName =
                    (string.IsNullOrWhiteSpace(Tag) ? "" : Tag + "_")
                    + LogDate + ".log";

                return Path.Combine(LogPath, fileName);
            }
        }

        public string LogDate
        {
            get { return DateTime.Now.ToString("yyyyMMdd"); }
        }

        public string preDate = DateTime.Now.ToString("yyyyMMdd");

        public string LineHeader
        {
            get { return $"[{DateTime.Now:HH:mm:ss}] "; }
        }

        // 기본 생성자
        public SoborLog()
        {
            StartLoop();
        }

        // 로그 경로 및 태그 지정
        public SoborLog(string path, string tag = "")
        {
            if (!string.IsNullOrWhiteSpace(path))
            {
                PathBase = NormalizePath(path);
            }

            Tag = tag ?? "";

            StartLoop();
        }

        // 경로 정규화
        private static string NormalizePath(string path)
        {
            string normalized = path.Trim();

            // 기존 경로 지정 방식과의 호환
            if (normalized.StartsWith("./", StringComparison.Ordinal))
            {
                normalized = @"C:\Soborsoft\log\";
            }

            // Windows 경로 구분자 정리
            normalized = normalized.Replace('/', Path.DirectorySeparatorChar);

            return Path.GetFullPath(normalized);
        }

        // 로그 루프 시작
        public void Start()
        {
            StartLoop();
        }

        public void StartLoop()
        {
            lock (_loopLock)
            {
                if (_disposed)
                    throw new ObjectDisposedException(nameof(SoborLog));

                if (_loopTask != null && !_loopTask.IsCompleted)
                    return;

                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = new CancellationTokenSource();

                _loopTask = LoopLogAsync(_cancellationTokenSource.Token);
            }
        }

        // 로그 루프 종료 요청
        public void StopLoop()
        {
            lock (_loopLock)
            {
                _cancellationTokenSource?.Cancel();
            }
        }

        // 로그 루프 종료 대기
        public async Task StopLoopAsync()
        {
            Task loopTask;

            lock (_loopLock)
            {
                _cancellationTokenSource?.Cancel();
                loopTask = _loopTask;
            }

            if (loopTask != null)
            {
                try
                {
                    await loopTask.ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    // 정상적인 취소
                }
            }
        }

        // 로그 처리 루프
        private async Task LoopLogAsync(CancellationToken token)
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    if (DoSave)
                    {
                        WriteLogList();
                    }

                    if (DoDeleteOldLog)
                    {
                        DeleteOldLog();
                    }

                    int delay = Math.Max(100, Interval);

                    await Task.Delay(delay, token).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException)
            {
                // 종료 요청
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"SoborLog loop error: {ex}");
            }
            finally
            {
                // 종료 직전 남아 있는 로그도 저장
                try
                {
                    WriteLogList();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"SoborLog final write error: {ex}");
                }
            }
        }

        // 로그 추가
        public void Add(string str, bool useHeader = true)
        {
            if (_disposed)
                return;

            string message = str ?? "";

            bool addHeader = useHeader && UseLineHeader;

            if (addHeader)
            {
                message = LineHeader + message;
            }

            logList.Enqueue(message);
        }

        // 여러 로그 추가
        public void AddRange(IEnumerable<string> messages, bool useHeader = true)
        {
            if (messages == null)
                return;

            foreach (string message in messages)
            {
                Add(message, useHeader);
            }
        }

        // 대기 중인 로그를 파일에 기록
        private void WriteLogList()
        {
            List<string> logs = GetLogList();

            if (logs == null || logs.Count == 0)
                return;

            try
            {
                Directory.CreateDirectory(LogPath);

                using (StreamWriter sw = new StreamWriter(
                    LogFile,
                    true,
                    Encoding.Default))
                {
                    foreach (string log in logs)
                    {
                        if (string.IsNullOrEmpty(log))
                            continue;

                        sw.WriteLine(log);

#if DEBUG
                        System.Diagnostics.Debug.WriteLine(log);
#endif
                    }
                }

                // 파일 기록 성공 후 화면에 출력
                foreach (string log in logs)
                {
                    if (!string.IsNullOrEmpty(log))
                    {
                        WriteLogToListView(log);
                    }
                }
            }
            catch (Exception ex)
            {
                // 파일 쓰기 실패 시 로그 유실을 줄이기 위해 큐에 복원
                for (int i = logs.Count - 1; i >= 0; i--)
                {
                    logList.Enqueue(logs[i]);
                }

                System.Diagnostics.Debug.WriteLine(
                    $"SoborLog write error: {ex}");
            }
        }

        // WPF ListBox에 로그 출력
        private void WriteLogToListView(string log)
        {
            if (allLog == null)
                return;

            if (!allLog.Dispatcher.CheckAccess())
            {
                if (!allLog.Dispatcher.HasShutdownStarted &&
                    !allLog.Dispatcher.HasShutdownFinished)
                {
                    allLog.Dispatcher.BeginInvoke(
                        new Action(() => AppendLog(allLog, log)),
                        DispatcherPriority.Background);
                }

                return;
            }

            AppendLog(allLog, log);
        }

        // 화면 로그 추가 및 최대 줄 수 유지
        private static void AppendLog(ListBox listBox, string log)
        {
            if (listBox == null)
                return;

            listBox.Items.Add(log);

            while (listBox.Items.Count > MaxLogLines)
            {
                listBox.Items.RemoveAt(0);
            }

            if (listBox.Items.Count > 0)
            {
                listBox.ScrollIntoView(
                    listBox.Items[listBox.Items.Count - 1]);
            }
        }

        // 로그 큐에서 목록 가져오기
        public List<string?> GetLogList()
        {
            List<string?> logs = new List<string?>();

            while (logList.TryDequeue(out string? log))
            {
                logs.Add(log);
            }

            return logs.Count == 0 ? null : logs;
        }

        // 오래된 로그 파일 삭제
        private void DeleteOldLog()
        {
            string today = LogDate;

            // 날짜가 바뀐 경우 하루에 한 번 정리
            if (_lastCleanupDate == today)
                return;

            _lastCleanupDate = today;
            preDate = today;

            try
            {
                if (!Directory.Exists(PathBase))
                    return;

                DateTime cutoffDate = DateTime.Now.Date.AddDays(-Math.Max(0, Period));

                // 연도/월별 하위 폴더의 로그 파일 검색
                foreach (string file in Directory.EnumerateFiles(
                    PathBase,
                    "*.log",
                    SearchOption.AllDirectories))
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);

                        if (!TryGetLogDate(fileName, out DateTime logDate))
                            continue;

                        if (logDate.Date < cutoffDate)
                        {
                            File.Delete(file);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"SoborLog delete error: {file}, {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"SoborLog cleanup error: {ex}");
            }
        }

        // 파일명 끝의 yyyyMMdd 날짜 추출
        // 예: 20261015.log, MyTag_20261015.log
        private static bool TryGetLogDate(
            string fileName,
            out DateTime logDate)
        {
            logDate = DateTime.MinValue;

            if (string.IsNullOrEmpty(fileName) || fileName.Length < 8)
                return false;

            string dateText = fileName.Substring(fileName.Length - 8, 8);

            return DateTime.TryParseExact(
                dateText,
                "yyyyMMdd",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out logDate);
        }

        // 자원 해제
        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;

            try
            {
                StopLoopAsync().GetAwaiter().GetResult();
            }
            catch
            {
                // 종료 중 예외 무시
            }

            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;

            GC.SuppressFinalize(this);
        }
    }
}
