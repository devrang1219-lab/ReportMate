using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media.Imaging;

namespace SoborSmartReport.SortImage
{
    public class ImageItem : INotifyPropertyChanged
    {
        private int _order;
        private int _thumbnailWidth = 120;
        private int _thumbnailHeight = 120;
        private BitmapSource _thumbnail;

        public int Order
        {
            get => _order;
            set
            {
                _order = value;
                OnPropertyChanged();
            }
        }

        public string FullPath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string FileSize { get; set; } = string.Empty;
        public DateTime LastWriteTime { get; set; }

        public BitmapSource Thumbnail
        {
            get => _thumbnail;
            set
            {
                _thumbnail = value;
                OnPropertyChanged();
            }
        }

        public int ThumbnailWidth
        {
            get => _thumbnailWidth;
            set
            {
                _thumbnailWidth = value;
                OnPropertyChanged();
            }
        }

        public int ThumbnailHeight
        {
            get => _thumbnailHeight;
            set
            {
                _thumbnailHeight = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(
            [CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(
                this,
                new PropertyChangedEventArgs(propertyName));
        }
    }
}