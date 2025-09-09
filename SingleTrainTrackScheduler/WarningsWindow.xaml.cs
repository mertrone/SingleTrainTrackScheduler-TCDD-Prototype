using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace SingleTrainTrackScheduler
{
    public partial class WarningsWindow : Window
    {
        public WarningsWindow()
        {
            InitializeComponent();
        }

        // MainWindow bu event’lere abone olacak
        public event Action AcceptAllRequested;
        public event Action<string> AcceptSelectedRequested;

        // MainWindow buradan satırları basıyor
        public void SetItems(IEnumerable<string> lines)
        {
            lstWarn.ItemsSource = (lines ?? Enumerable.Empty<string>()).ToList();
        }

        private void BtnAcceptAll_Click(object sender, RoutedEventArgs e)
        {
            AcceptAllRequested?.Invoke();
        }

        private void BtnAcceptSelected_Click(object sender, RoutedEventArgs e)
        {
            var line = lstWarn.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(line))
            {
                MessageBox.Show("Önce listeden bir tren seçin.");
                return;
            }

            // Satırlar şu formatta: "FU01 [P] +20 dk (> 15 dk) — feasible, OPT değil"
            // -> ilk token tren id
            var tid = line.Split(' ').FirstOrDefault()?.Trim();
            if (!string.IsNullOrWhiteSpace(tid))
                AcceptSelectedRequested?.Invoke(tid);
        }
    }
}
