using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using CybageMISAutomation.Models;
using CybageMISAutomation.Services;
using System;

namespace CybageMISAutomation
{
    public partial class ComparisonWindow : Window
    {
        public ObservableCollection<SwipeLogEntry> YesterdayEntries { get; set; }
        public ObservableCollection<SwipeLogEntry> TodayEntries { get; set; }
        private MainWindow? _mainWindow;

        public ComparisonWindow(MainWindow? mainWindow = null)
        {
            InitializeComponent();
            _mainWindow = mainWindow;
            
            YesterdayEntries = new ObservableCollection<SwipeLogEntry>();
            TodayEntries = new ObservableCollection<SwipeLogEntry>();
            
            DataGridYesterday.ItemsSource = YesterdayEntries;
            DataGridToday.ItemsSource = TodayEntries;
        }

        public void LoadComparisonData(List<SwipeLogEntry> yesterdayData, List<SwipeLogEntry> todayData)
        {
            var now = DateTime.Now;
            var yesterday = WorkHoursCalculator.Calculate(yesterdayData, now);
            var today = WorkHoursCalculator.Calculate(todayData, now);

            YesterdayEntries.Clear();
            foreach (var entry in yesterday.Labeled) YesterdayEntries.Add(entry);

            TodayEntries.Clear();
            foreach (var entry in today.Labeled) TodayEntries.Add(entry);

            TxtYesterdayCount.Text = $"Yesterday: {yesterdayData.Count} records" + SkippedNote(yesterday);
            TxtTodayCount.Text = $"Today: {todayData.Count} records" + SkippedNote(today);

            TxtYesterdayWorkHours.Text = $"Yesterday Work Hours: {yesterday.WorkTimeDisplay}";
            TxtTodayWorkHours.Text = $"Today Work Hours: {today.WorkTimeDisplay}" +
                (today.OngoingTime > TimeSpan.Zero ? $" (incl. {WorkHoursCalculator.Format(today.OngoingTime)} ongoing)" : "");
        }

        private static string SkippedNote(WorkHoursCalculator.Result r) =>
            r.SkippedRows.Count > 0 ? $" ({r.SkippedRows.Count} unreadable row(s) skipped)" : "";

        private void BtnExportBoth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    DefaultExt = "csv",
                    FileName = $"SwipeLog_Comparison_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var csv = new StringBuilder();
                    
                    // Header
                    csv.AppendLine("Report Type,Employee ID,Date,Machine Name,Direction,Time,Activity,Duration");
                    
                    // Yesterday data
                    foreach (var entry in YesterdayEntries)
                    {
                        csv.AppendLine($"Yesterday,\"{entry.EmployeeId}\",\"{entry.Date}\",\"{entry.MachineName}\",\"{entry.Direction}\",\"{entry.Time}\",\"{entry.Activity}\",\"{entry.Duration}\"");
                    }
                    
                    // Today data
                    foreach (var entry in TodayEntries)
                    {
                        csv.AppendLine($"Today,\"{entry.EmployeeId}\",\"{entry.Date}\",\"{entry.MachineName}\",\"{entry.Direction}\",\"{entry.Time}\",\"{entry.Activity}\",\"{entry.Duration}\"");
                    }
                    
                    File.WriteAllText(saveDialog.FileName, csv.ToString());
                    MessageBox.Show($"Comparison data exported successfully to:\n{saveDialog.FileName}", 
                                    "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting data: {ex.Message}", "Export Error", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnCopyBoth_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var text = new StringBuilder();
                
                // Header
                text.AppendLine("Report Type\tEmployee ID\tDate\tMachine Name\tDirection\tTime\tActivity\tDuration");
                
                // Yesterday data
                foreach (var entry in YesterdayEntries)
                {
                    text.AppendLine($"Yesterday\t{entry.EmployeeId}\t{entry.Date}\t{entry.MachineName}\t{entry.Direction}\t{entry.Time}\t{entry.Activity}\t{entry.Duration}");
                }
                
                // Today data
                foreach (var entry in TodayEntries)
                {
                    text.AppendLine($"Today\t{entry.EmployeeId}\t{entry.Date}\t{entry.MachineName}\t{entry.Direction}\t{entry.Time}\t{entry.Activity}\t{entry.Duration}");
                }
                
                Clipboard.SetText(text.ToString());
                MessageBox.Show("Comparison data copied to clipboard!", "Copy Complete", 
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying data: {ex.Message}", "Copy Error", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnRunAgain_Click(object sender, RoutedEventArgs e)
        {
            if (_mainWindow != null)
            {
                await _mainWindow.StartFullAutomation();
                this.Close();
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}