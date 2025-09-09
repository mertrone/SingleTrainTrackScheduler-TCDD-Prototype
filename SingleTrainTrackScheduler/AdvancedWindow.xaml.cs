using Microsoft.Win32;
using System.ComponentModel;
using System.Linq;
using System.Windows;

namespace SingleTrainTrackScheduler
{
    public partial class AdvancedWindow : Window
    {
        private readonly MainWindow _main;
        public int Head => TryGet(txtHead.Text, 3);
        public int Clear => TryGet(txtClear.Text, 3);
        public int Slack => TryGet(txtSlack.Text, 1);
        protected override void OnClosing(CancelEventArgs e)
        {
            // Pencereyi gerçekten kapatmak yerine gizle
            e.Cancel = true;
            this.Hide();
            base.OnClosing(e);
        }
        public AdvancedWindow(MainWindow owner)
        {
            InitializeComponent();
            _main = owner;

            // UI'yi mevcut değerlerle doldur
            txtHead.Text = _main?.GetHead().ToString() ?? txtHead.Text;
            txtClear.Text = _main?.GetClear().ToString() ?? txtClear.Text;
            txtSlack.Text = _main?.GetSlack().ToString() ?? txtSlack.Text;
        }

        public void SyncStateFromMain()
        {
            lstStations.Items.Clear();
            foreach (var name in _main?.StationNames() ?? Enumerable.Empty<string>())
                lstStations.Items.Add(name);

            txtConfDialog.Text = _main?.ConflictSummaryText() ?? "";
        }

        private void BtnCheck_Click(object sender, RoutedEventArgs e) => _main?.InvokeCheckConflicts();
        private void BtnOptimize_Click(object s, RoutedEventArgs e) => _main?.InvokeOptimize();
        private void BtnApplyBuffers_Click(object sender, RoutedEventArgs e)
        {
            // TextBox'lardaki Head/Clear/Slack değerleri zaten Head/Clear/Slack getter'larından okunuyor.
            // Sadece yeniden çiz ve özet/uyarıyı güncelle.
            _main?.InvokeRefreshPlot();
            _main?.InvokeCheckConflicts();
        }

        private void BtnExportCsv_Click(object sender, RoutedEventArgs e)
        {
            var sfd = new SaveFileDialog { Filter = "CSV|*.csv", FileName = "timetable.csv" };
            if (sfd.ShowDialog(this) != true) return;
            var only = lstStations.SelectedItems.Cast<string>().ToList();
            _main?.ExportTimetableCsv(sfd.FileName, only, true, true);
            MessageBox.Show("Kaydedildi.", "CSV");
        }

        private void BtnTemplate_Click(object sender, RoutedEventArgs e)
        {
            var sfd = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = "dataset_template.xlsx" };
            if (sfd.ShowDialog(this) != true) return;
            _main?.SaveExcelTemplate(sfd.FileName);
        }

        private void BtnSaveXlsx_Click(object sender, RoutedEventArgs e)
        {
            var sfd = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = "dataset.xlsx" };
            if (sfd.ShowDialog(this) != true) return;
            _main?.SaveExcelDataset(sfd.FileName);
        }

        private void BtnLoadXlsx_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog { Filter = "Excel|*.xlsx;*.xls" };
            if (ofd.ShowDialog(this) != true) return;
            _main?.LoadExcelDataset(ofd.FileName);
            SyncStateFromMain();
        }

        private static int TryGet(string s, int def) => int.TryParse(s, out var v) ? v : def;
    }
}
