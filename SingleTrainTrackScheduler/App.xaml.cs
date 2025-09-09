using SingleTrainTrackScheduler.Interop;
using System;
using System.IO;
using System.Windows;

namespace SingleTrainTrackScheduler
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                NativeOrtools.EnsureLoaded();
            }
            catch (Exception ex)
            {
                var log = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "SingleTrainTrackScheduler_startup_error.txt");
                File.WriteAllText(log, ex.ToString());
                MessageBox.Show("Başlangıç hatası:\n\n" + ex.Message +
                                "\n\nAyrıntılar masaüstü log dosyasında.",
                                "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown(1);
                return;
            }
            base.OnStartup(e);
        }
    }
}
