
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace SoborSmartReport.SortImage
{
    public partial class FormSortImage : UserControl
    {
        private readonly ObservableCollection<ImageItem> _images = new();

        private string _currentFolder = string.Empty;
        private Point _dragStartPoint;
        private ImageItem _draggedItem;
        private bool _isInitializing;
        private bool _isUpdatingSort;

        public FormSortImage()
        {
            InitializeComponent();

            _isInitializing = true;

            lvImages.ItemsSource = _images;

            _isInitializing = false;

            Loaded += FormSortImage_Loaded;
        }

        // =========================================================
        // 초기화 및 폴더 트리
        // =========================================================

        private void FormSortImage_Loaded(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tbFolder.Text))
            {
                string pictures = Environment.GetFolderPath(
                    Environment.SpecialFolder.MyPictures);

                if (Directory.Exists(pictures))
                {
                    tbFolder.Text = pictures;
                    LoadFolderTree(pictures);
                    LoadImages(pictures);
                }
            }
        }

        private void LoadFolderTree(string path)
        {
            tvFolders.Items.Clear();

            if (!Directory.Exists(path))
                return;

            var root = new TreeViewItem
            {
                Header = Path.GetFileName(path.TrimEnd(
                    Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar)),
                Tag = path,
                IsExpanded = true
            };

            AddSubFolders(root, path);

            tvFolders.Items.Add(root);
            root.IsSelected = true;
        }

        private void AddSubFolders(TreeViewItem parent, string path)
        {
            try
            {
                foreach (string directory in Directory.GetDirectories(path))
                {
                    var item = new TreeViewItem
                    {
                        Header = Path.GetFileName(directory),
                        Tag = directory
                    };

                    parent.Items.Add(item);

                    // 하위 폴더는 선택할 때 로드하도록 지연합니다.
                    item.Items.Add(new TreeViewItem { Header = "..." });
                }
            }
            catch
            {
                // 접근 권한이 없는 폴더는 건너뜁니다.
            }
        }

        private void TvFolders_SelectedItemChanged(
            object sender,
            RoutedPropertyChangedEventArgs<object> e)
        {
            if (tvFolders.SelectedItem is not TreeViewItem item)
                return;

            if (item.Tag is not string path || !Directory.Exists(path))
                return;

            // 하위 폴더를 실제로 펼칠 때 로드
            if (item.Items.Count == 1 &&
                item.Items[0] is TreeViewItem placeholder &&
                Equals(placeholder.Header, "..."))
            {
                item.Items.Clear();
                AddSubFolders(item, path);
            }

            tbFolder.Text = path;
            LoadImages(path);
        }

        // =========================================================
        // 폴더 입력 및 선택
        // =========================================================

        private void TbFolder_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;

            string path = tbFolder.Text.Trim();

            if (!Directory.Exists(path))
            {
                MessageBox.Show(
                    "존재하는 폴더 경로를 입력해 주세요.",
                    "폴더 확인",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }

            LoadFolderTree(path);
            LoadImages(path);

            e.Handled = true;
        }

        private void BtnFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog
            {
                Title = "이미지 폴더 선택",
                Multiselect = false
            };

            if (dialog.ShowDialog() != true)
                return;

            tbFolder.Text = dialog.FolderName;
            LoadFolderTree(dialog.FolderName);
            LoadImages(dialog.FolderName);
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            string path = tbFolder.Text.Trim();

            if (!Directory.Exists(path))
            {
                MessageBox.Show(
                    "먼저 유효한 폴더를 선택하세요.",
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return;
            }

            LoadFolderTree(path);
            LoadImages(path);
        }

        // =========================================================
        // 이미지 로드
        // =========================================================

        private static bool IsImageFile(string path)
        {
            string ext = Path.GetExtension(path);

            return new[]
            {
                ".jpg", ".jpeg", ".png", ".bmp",
                ".gif", ".tif", ".tiff", ".webp"
            }.Contains(ext, StringComparer.OrdinalIgnoreCase);
        }

        private void LoadImages(string folder)
        {
            if (!Directory.Exists(folder))
                return;

            _currentFolder = folder;
            _images.Clear();

            try
            {
                var files = Directory
                    .EnumerateFiles(folder)
                    .Where(IsImageFile)
                    .ToList();

                IEnumerable<string> sorted = ApplySort(files);

                int order = 1;
                int thumbSize = GetThumbnailSize();

                foreach (string file in sorted)
                {
                    try
                    {
                        var item = new ImageItem
                        {
                            Order = order++,
                            FullPath = file,
                            FileName = Path.GetFileName(file),
                            LastWriteTime = File.GetLastWriteTime(file),
                            ThumbnailWidth = thumbSize,
                            ThumbnailHeight = thumbSize,
                            Thumbnail = LoadThumbnail(file, thumbSize)
                        };

                        _images.Add(item);
                    }
                    catch
                    {
                        // 읽을 수 없는 이미지 파일은 건너뜁니다.
                    }
                }

                if (_images.Count > 0)
                    lvImages.SelectedIndex = 0;
                else
                    ClearPreview();

                UpdateOrderNumbers();
                UpdatePreview();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"이미지 목록을 읽을 수 없습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private static BitmapImage LoadThumbnail(string path, int size)
        {
            var bitmap = new BitmapImage();

            using (var stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete))
            {
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.DecodePixelWidth = size;
                bitmap.DecodePixelHeight = size;
                bitmap.StreamSource = stream;
                bitmap.EndInit();
            }

            bitmap.Freeze();
            return bitmap;
        }

        public static BitmapImage LoadImageWithoutLock(string path)
        {
            var bitmap = new BitmapImage();

            using (var stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete))
            {
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.StreamSource = stream;
                bitmap.EndInit();
            }

            bitmap.Freeze();
            return bitmap;
        }

        private int GetThumbnailSize()
        {
            return Math.Clamp(
                (int)tbThumbSize.Value,
                80,
                250);
        }

        // =========================================================
        // 정렬
        // =========================================================

        private IEnumerable<string> ApplySort(IEnumerable<string> files)
        {
            int index = cbSort.SelectedIndex;

            switch (index)
            {
                case 1:
                    return files.OrderBy(
                        Path.GetFileName,
                        StringComparer.CurrentCultureIgnoreCase);

                case 2:
                    return files.OrderByDescending(
                        Path.GetFileName,
                        StringComparer.CurrentCultureIgnoreCase);

                case 3:
                    return files.OrderBy(File.GetLastWriteTime);

                case 4:
                    return files.OrderByDescending(File.GetLastWriteTime);

                default:
                    return files.OrderBy(
                        Path.GetFileName,
                        StringComparer.CurrentCultureIgnoreCase);
            }
        }

        private void CbSort_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            if (_isInitializing || _isUpdatingSort)
                return;

            if (!IsLoaded || !Directory.Exists(_currentFolder))
                return;

            // 사용자 지정 순서는 사용자가 직접 조정한 상태이므로
            // 정렬 콤보박스를 변경했을 때만 자동 정렬합니다.
            if (cbSort.SelectedIndex != 0)
                LoadImages(_currentFolder);
        }

        // =========================================================
        // 선택 및 미리보기
        // =========================================================

        private void LvImages_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void LvImages_MouseDoubleClick(
            object sender,
            MouseButtonEventArgs e)
        {
            OpenLargePreview();
        }

        private void PbPreview_MouseDown(
            object sender,
            MouseButtonEventArgs e)
        {
            if (e.ClickCount >= 2)
                OpenLargePreview();
        }

        private void BtnPrev_Click(object sender, RoutedEventArgs e)
        {
            if (lvImages.SelectedIndex > 0)
                lvImages.SelectedIndex--;
        }

        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            if (lvImages.SelectedIndex < _images.Count - 1)
                lvImages.SelectedIndex++;
        }

        private void UpdatePreview()
        {
            if (lvImages.SelectedItem is not ImageItem item)
            {
                ClearPreview();
                return;
            }

            try
            {
                pbPreview.Source = LoadImageWithoutLock(item.FullPath);
                lbFileName.Text = item.FileName;
                lbInfo.Text =
                    $"{item.Order:D3} / {_images.Count}    " +
                    $"{item.LastWriteTime:yyyy-MM-dd HH:mm}";
            }
            catch (Exception ex)
            {
                ClearPreview();

                lbFileName.Text = item.FileName;
                lbInfo.Text = "이미지를 표시할 수 없습니다.";

                Debug.WriteLine(ex);
            }
        }

        private void ClearPreview()
        {
            pbPreview.Source = null;
            lbFileName.Text = "이미지를 선택하세요.";
            lbInfo.Text = string.Empty;
        }

        private void OpenLargePreview()
        {
            if (lvImages.SelectedItem is not ImageItem selected)
                return;

            int index = _images.IndexOf(selected);

            if (index < 0)
                return;

            var preview = new PreviewWindow(_images.ToList(), index)
            {
                Owner = Window.GetWindow(this)
            };

            preview.ShowDialog();
        }

        // =========================================================
        // 위로 / 아래로 이동
        // =========================================================

        private void BtnUp_Click(object sender, RoutedEventArgs e)
        {
            MoveSelectedItems(-1);
        }

        private void BtnDown_Click(object sender, RoutedEventArgs e)
        {
            MoveSelectedItems(1);
        }

        private void MoveSelectedItems(int direction)
        {
            var selected = lvImages.SelectedItems
                .Cast<ImageItem>()
                .OrderBy(x => _images.IndexOf(x))
                .ToList();

            if (selected.Count == 0)
                return;

            if (direction < 0)
            {
                foreach (var item in selected)
                {
                    int index = _images.IndexOf(item);

                    if (index > 0)
                        _images.Move(index, index - 1);
                }
            }
            else
            {
                foreach (var item in selected.OrderByDescending(
                    x => _images.IndexOf(x)))
                {
                    int index = _images.IndexOf(item);

                    if (index < _images.Count - 1)
                        _images.Move(index, index + 1);
                }
            }

            UpdateOrderNumbers();
            RestoreSelection(selected);
            cbSort.SelectedIndex = 0;
        }

        // =========================================================
        // 목록 제외 - 원본 파일은 삭제하지 않음
        // =========================================================

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            var selected = lvImages.SelectedItems
                .Cast<ImageItem>()
                .ToList();

            if (selected.Count == 0)
                return;

            foreach (var item in selected)
                _images.Remove(item);

            UpdateOrderNumbers();
            UpdatePreview();
        }

        // =========================================================
        // 드래그 앤 드롭 정렬
        // =========================================================

        private void LvImages_PreviewMouseLeftButtonDown(
            object sender,
            MouseButtonEventArgs e)
        {
            _dragStartPoint = e.GetPosition(lvImages);

            if (e.OriginalSource is DependencyObject source)
            {
                var container = FindAncestor<ListBoxItem>(source);
                _draggedItem = container?.DataContext as ImageItem;
            }
        }

        private void LvImages_PreviewMouseMove(
            object sender,
            MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed ||
                _draggedItem == null)
                return;

            Point current = e.GetPosition(lvImages);

            if (Math.Abs(current.X - _dragStartPoint.X) <
                    SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(current.Y - _dragStartPoint.Y) <
                    SystemParameters.MinimumVerticalDragDistance)
                return;

            DragDrop.DoDragDrop(
                lvImages,
                _draggedItem,
                DragDropEffects.Move);

            _draggedItem = null;
        }

        private void LvImages_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(typeof(ImageItem))
                ? DragDropEffects.Move
                : DragDropEffects.None;

            e.Handled = true;
        }

        private void LvImages_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(ImageItem)))
                return;

            if (e.Data.GetData(typeof(ImageItem)) is not ImageItem source)
                return;

            if (e.OriginalSource is not DependencyObject targetObject)
                return;

            var targetContainer = FindAncestor<ListBoxItem>(targetObject);

            if (targetContainer?.DataContext is not ImageItem target)
                return;

            int sourceIndex = _images.IndexOf(source);
            int targetIndex = _images.IndexOf(target);

            if (sourceIndex < 0 || targetIndex < 0 ||
                sourceIndex == targetIndex)
                return;

            _images.Move(sourceIndex, targetIndex);

            UpdateOrderNumbers();
            RestoreSelection(new[] { source });
            cbSort.SelectedIndex = 0;
        }

        private static T FindAncestor<T>(DependencyObject current)
            where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T result)
                    return result;

                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }

        // =========================================================
        // 썸네일 크기 변경
        // =========================================================

        private void TbThumbSize_ValueChanged(
            object sender,
            RoutedPropertyChangedEventArgs<double> e)
        {
            if (!IsLoaded || _images.Count == 0)
                return;

            int size = Math.Clamp((int)e.NewValue, 80, 250);

            foreach (var item in _images)
            {
                item.ThumbnailWidth = size;
                item.ThumbnailHeight = size;

                try
                {
                    item.Thumbnail = LoadThumbnail(item.FullPath, size);
                }
                catch
                {
                    // 읽을 수 없는 파일은 그대로 둡니다.
                }
            }
        }

        // =========================================================
        // 순번 갱신 및 선택 복원
        // =========================================================

        private void UpdateOrderNumbers()
        {
            for (int i = 0; i < _images.Count; i++)
                _images[i].Order = i + 1;

            UpdatePreview();
        }

        private void RestoreSelection(IEnumerable<ImageItem> items)
        {
            lvImages.SelectedItems.Clear();

            foreach (var item in items)
            {
                if (_images.Contains(item))
                    lvImages.SelectedItems.Add(item);
            }

            if (lvImages.SelectedItem is ImageItem selected)
                lvImages.ScrollIntoView(selected);
        }

        // =========================================================
        // 정렬본 만들기
        // =========================================================

        private void BtnMakeFolder_Click(object sender, RoutedEventArgs e)
        {
            if (_images.Count == 0)
            {
                MessageBox.Show(
                    "저장할 이미지가 없습니다.",
                    "알림",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return;
            }

            var dialog = new OpenFolderDialog
            {
                Title = "정렬본을 저장할 상위 폴더 선택",
                Multiselect = false
            };

            if (dialog.ShowDialog() != true)
                return;

            string outputFolder = Path.Combine(
                dialog.FolderName,
                "정렬본_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));

            try
            {
                Directory.CreateDirectory(outputFolder);

                for (int i = 0; i < _images.Count; i++)
                {
                    ImageItem item = _images[i];
                    string extension = Path.GetExtension(item.FullPath);

                    string fileName;

                    if (chkKeepOriginalName.IsChecked == true)
                    {
                        string originalName = Path.GetFileNameWithoutExtension(
                            item.FileName);

                        fileName = $"{i + 1:D3}_{originalName}{extension}";
                    }
                    else
                    {
                        fileName = $"{i + 1:D3}{extension}";
                    }

                    string destination = Path.Combine(
                        outputFolder,
                        fileName);

                    File.Copy(item.FullPath, destination, false);
                }

                MessageBox.Show(
                    $"정렬본을 만들었습니다.\n\n{outputFolder}",
                    "완료",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                Process.Start(new ProcessStartInfo
                {
                    FileName = outputFolder,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"정렬본 생성 중 오류가 발생했습니다.\n{ex.Message}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }
}
