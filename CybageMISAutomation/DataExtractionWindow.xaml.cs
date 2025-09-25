using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using CybageMISAutomation.Models;

namespace CybageMISAutomation
{
    public partial class DataExtractionWindow : Window
    {
        public ObservableCollection<SwipeLogEntry> SwipeLogEntries { get; set; }

        public DataExtractionWindow()
        {
            InitializeComponent();
            SwipeLogEntries = new ObservableCollection<SwipeLogEntry>();
            DataGridResults.ItemsSource = SwipeLogEntries;
        }

        public void LoadSwipeLogData(List<SwipeLogEntry> entries, string reportType)
        {
            SwipeLogEntries.Clear();
            foreach (var entry in entries)
            {
                SwipeLogEntries.Add(entry);
            }
            
            TxtReportType.Text = $"({reportType} Report)";
            TxtRecordCount.Text = $"Records: {entries.Count}";
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    DefaultExt = "csv",
                    FileName = $"SwipeLog_{TxtReportType.Text.Replace("(", "").Replace(" Report)", "")}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var csv = new StringBuilder();
                    csv.AppendLine("Employee ID,Employee Name,Date,In Time,Out Time,Duration,Status,Location");
                    
                    foreach (var entry in SwipeLogEntries)
                    {
                        csv.AppendLine($"\"{entry.EmployeeId}\",\"{entry.EmployeeName}\",\"{entry.Date}\",\"{entry.InTime}\",\"{entry.OutTime}\",\"{entry.Duration}\",\"{entry.Status}\",\"{entry.Location}\"");
                    }
                    
                    File.WriteAllText(saveDialog.FileName, csv.ToString());
                    MessageBox.Show($"Data exported successfully to:\n{saveDialog.FileName}", "Export Complete", 
                                    MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting data: {ex.Message}", "Export Error", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var text = new StringBuilder();
                text.AppendLine("Employee ID\tEmployee Name\tDate\tIn Time\tOut Time\tDuration\tStatus\tLocation");
                
                foreach (var entry in SwipeLogEntries)
                {
                    text.AppendLine($"{entry.EmployeeId}\t{entry.EmployeeName}\t{entry.Date}\t{entry.InTime}\t{entry.OutTime}\t{entry.Duration}\t{entry.Status}\t{entry.Location}");
                }
                
                Clipboard.SetText(text.ToString());
                MessageBox.Show("Data copied to clipboard!", "Copy Complete", 
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying data: {ex.Message}", "Copy Error", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}