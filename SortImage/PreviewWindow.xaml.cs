using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace SoborSmartReport.SortImage
{
    public partial class PreviewWindow : Window
    {
        private readonly List<ImageItem> _images;
        private int _index;

        // 0이면 화면 맞춤, 1이면 원본 크기 기준 100%
        private double _zoom = 0;

        private BitmapSource _currentBitmap;

        public PreviewWindow(
            List<ImageItem> images,
            int startIndex)
        {
            InitializeComponent();

            _images = images
                ?? throw new ArgumentNullException(nameof(images));

            if (_images.Count == 0)
                throw new ArgumentException(
                    "미리보기 이미지가 없습니다.",
                    nameof(images));

            _index = Math.Clamp(
                startIndex, 0, _images.Count - 1);

            Loaded += (_, _) => ShowImage();

            SizeChanged += (_, _) =>
            {
                if (_zoom == 0)
                    FitImage();
            };
        }

        // 이미지 파일 잠금 방지를 위해 메모리로 로드
        private static BitmapImage LoadBitmap(string path)
        {
            var bitmap = new BitmapImage();

            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.UriSource = new Uri(
                System.IO.Path.GetFullPath(path),
                UriKind.Absolute);
            bitmap.EndInit();

            bitmap.Freeze();

            return bitmap;
        }

        private void ShowImage()
        {
            if (_images.Count == 0)
                return;

            _index = Math.Clamp(
                _index, 0, _images.Count - 1);

            ImageItem item = _images[_index];

            try
            {
                _currentBitmap = LoadBitmap(item.FullPath);
                pictureBox.Source = _currentBitmap;

                Title = $"{item.Order:D3} - {item.FileName}";
                lbInfo.Text =
                    $"{_index + 1} / {_images.Count}    {item.FileName}";

                _zoom = 0;
                FitImage();
            }
            catch (Exception ex)
            {
                pictureBox.Source = null;
                _currentBitmap = null;

                MessageBox.Show(
                    this,
                    $"이미지를 열 수 없습니다.\n{item.FullPath}\n\n{ex.Message}",
                    "이미지 열기 오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void Previous()
        {
            if (_index <= 0)
                return;

            _index--;
            ShowImage();
        }

        private void Next()
        {
            if (_index >= _images.Count - 1)
                return;

            _index++;
            ShowImage();
        }

        private void FitImage()
        {
            if (_currentBitmap == null)
                return;

            double availableWidth = imageScroll.ViewportWidth;
            double availableHeight = imageScroll.ViewportHeight;

            if (availableWidth <= 0 || availableHeight <= 0)
                return;

            double scale = Math.Min(
                availableWidth / _currentBitmap.PixelWidth,
                availableHeight / _currentBitmap.PixelHeight);

            scale = Math.Max(0.01, scale);

            pictureBox.Width = _currentBitmap.PixelWidth * scale;
            pictureBox.Height = _currentBitmap.PixelHeight * scale;

            _zoom = 0;
            UpdateInfo();
        }

        private void Zoom100()
        {
            if (_currentBitmap == null)
                return;

            pictureBox.Width = _currentBitmap.PixelWidth;
            pictureBox.Height = _currentBitmap.PixelHeight;

            _zoom = 1.0;
            UpdateInfo();
        }

        private void ZoomBy(double factor)
        {
            if (_currentBitmap == null)
                return;

            // 화면 맞춤 상태에서 현재 표시 배율을 계산
            if (_zoom == 0)
            {
                _zoom = pictureBox.Width /
                        _currentBitmap.PixelWidth;
            }

            _zoom = Math.Clamp(_zoom * factor, 0.05, 8.0);

            pictureBox.Width =
                _currentBitmap.PixelWidth * _zoom;

            pictureBox.Height =
                _currentBitmap.PixelHeight * _zoom;

            // 확대 시 이미지가 왼쪽 위부터 스크롤되도록 설정
            pictureBox.HorizontalAlignment = HorizontalAlignment.Left;
            pictureBox.VerticalAlignment = VerticalAlignment.Top;

            UpdateInfo();
        }

        private void UpdateInfo()
        {
            if (_currentBitmap == null)
                return;

            string fileName = _images[_index].FileName;

            string zoomText = _zoom == 0
                ? "화면 맞춤"
                : $"{_zoom * 100:F0}%";

            lbInfo.Text =
                $"{_index + 1} / {_images.Count}    " +
                $"{zoomText}    {fileName}";
        }

        private void BtnPrev_Click(
            object sender, RoutedEventArgs e) => Previous();

        private void BtnNext_Click(
            object sender, RoutedEventArgs e) => Next();

        private void BtnFit_Click(
            object sender, RoutedEventArgs e)
        {
            pictureBox.HorizontalAlignment = HorizontalAlignment.Center;
            pictureBox.VerticalAlignment = VerticalAlignment.Center;
            FitImage();
        }

        private void Btn100_Click(
            object sender, RoutedEventArgs e) => Zoom100();

        private void PreviewWindow_PreviewMouseWheel(
            object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.Control)
                return;

            ZoomBy(e.Delta > 0 ? 1.15 : 1.0 / 1.15);
            e.Handled = true;
        }

        private void PreviewWindow_KeyDown(
            object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Left:
                    Previous();
                    e.Handled = true;
                    break;

                case Key.Right:
                    Next();
                    e.Handled = true;
                    break;

                case Key.Home:
                    _index = 0;
                    ShowImage();
                    e.Handled = true;
                    break;

                case Key.End:
                    _index = _images.Count - 1;
                    ShowImage();
                    e.Handled = true;
                    break;

                case Key.Escape:
                    Close();
                    e.Handled = true;
                    break;

                case Key.D0:
                case Key.NumPad0:
                    Zoom100();
                    e.Handled = true;
                    break;

                case Key.F:
                    pictureBox.HorizontalAlignment = HorizontalAlignment.Center;
                    pictureBox.VerticalAlignment = VerticalAlignment.Center;
                    FitImage();
                    e.Handled = true;
                    break;
            }
        }
    }
}