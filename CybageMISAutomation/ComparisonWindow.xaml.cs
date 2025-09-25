using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using CybageMISAutomation.Models;

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
            // Load Yesterday data
            YesterdayEntries.Clear();
            foreach (var entry in yesterdayData)
            {
                YesterdayEntries.Add(entry);
            }
            
            // Load Today data
            TodayEntries.Clear();
            foreach (var entry in todayData)
            {
                TodayEntries.Add(entry);
            }
            
            // Update counts
            TxtYesterdayCount.Text = $"Yesterday: {yesterdayData.Count} records";
            TxtTodayCount.Text = $"Today: {todayData.Count} records";
        }

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
                    csv.AppendLine("Report Type,Employee ID,Employee Name,Date,In Time,Out Time,Duration,Status,Location");
                    
                    // Yesterday data
                    foreach (var entry in YesterdayEntries)
                    {
                        csv.AppendLine($"Yesterday,\"{entry.EmployeeId}\",\"{entry.EmployeeName}\",\"{entry.Date}\",\"{entry.InTime}\",\"{entry.OutTime}\",\"{entry.Duration}\",\"{entry.Status}\",\"{entry.Location}\"");
                    }
                    
                    // Today data
                    foreach (var entry in TodayEntries)
                    {
                        csv.AppendLine($"Today,\"{entry.EmployeeId}\",\"{entry.EmployeeName}\",\"{entry.Date}\",\"{entry.InTime}\",\"{entry.OutTime}\",\"{entry.Duration}\",\"{entry.Status}\",\"{entry.Location}\"");
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
                text.AppendLine("Report Type\tEmployee ID\tEmployee Name\tDate\tIn Time\tOut Time\tDuration\tStatus\tLocation");
                
                // Yesterday data
                foreach (var entry in YesterdayEntries)
                {
                    text.AppendLine($"Yesterday\t{entry.EmployeeId}\t{entry.EmployeeName}\t{entry.Date}\t{entry.InTime}\t{entry.OutTime}\t{entry.Duration}\t{entry.Status}\t{entry.Location}");
                }
                
                // Today data
                foreach (var entry in TodayEntries)
                {
                    text.AppendLine($"Today\t{entry.EmployeeId}\t{entry.EmployeeName}\t{entry.Date}\t{entry.InTime}\t{entry.OutTime}\t{entry.Duration}\t{entry.Status}\t{entry.Location}");
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