using System.Configuration;
using System.Data;
using System.Windows;

namespace SoborSmartReport
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            DispatcherUnhandledException += (sender, args) =>
            {
                MessageBox.Show(
                    args.Exception.Message,
                    "소보르소프트 엑셀보고서 도우미",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                args.Handled = true;
            };
        }
    }

}
