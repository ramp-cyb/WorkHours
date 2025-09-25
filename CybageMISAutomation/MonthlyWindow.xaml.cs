using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using CybageMISAutomation.Models;

namespace CybageMISAutomation
{
    public partial class MonthlyWindow : Window
    {
        public ObservableCollection<MonthlyAttendanceEntry> MonthlyEntries { get; set; }
        private string CurrentEmployeeId { get; set; } = "";

        public MonthlyWindow()
        {
            InitializeComponent();
            MonthlyEntries = new ObservableCollection<MonthlyAttendanceEntry>();
            DataGridMonthly.ItemsSource = MonthlyEntries;
        }

        public MonthlyWindow(string employeeId) : this()
        {
            CurrentEmployeeId = employeeId;
        }

        public void LoadMonthlyData(List<MonthlyAttendanceEntry> entries, string employeeId)
        {
            MonthlyEntries.Clear();
            foreach (var entry in entries)
            {
                MonthlyEntries.Add(entry);
            }

            CurrentEmployeeId = employeeId;
            UpdateStatistics();
        }

        private void UpdateStatistics()
        {
            if (MonthlyEntries.Count == 0) return;

            var firstEntry = MonthlyEntries.FirstOrDefault();
            if (firstEntry != null)
            {
                TxtEmployeeName.Text = $"Employee: {firstEntry.EmployeeName} ({firstEntry.EmployeeId})";
            }

            var dates = MonthlyEntries.Where(e => !string.IsNullOrEmpty(e.Date)).Select(e => e.Date).ToList();
            if (dates.Count > 0)
            {
                TxtDateRange.Text = $"Period: {dates.First()} to {dates.Last()}";
            }

            TxtTotalDays.Text = $"Total Days: {MonthlyEntries.Count}";

            // Calculate average and total work hours
            var workHours = MonthlyEntries
                .Where(e => !string.IsNullOrEmpty(e.ActualWorkHours) && e.ActualWorkHours != "0:00")
                .Select(e => ParseWorkHours(e.ActualWorkHours))
                .Where(h => h.TotalMinutes > 0)
                .ToList();

            if (workHours.Count > 0)
            {
                var totalMinutes = workHours.Sum(h => (int)h.TotalMinutes);
                var avgMinutes = totalMinutes / workHours.Count;

                TxtTotalWorkHours.Text = $"Total Work Hours: {FormatMinutes(totalMinutes)}";
                TxtAvgWorkHours.Text = $"Average Work Hours: {FormatMinutes(avgMinutes)}";
            }
            else
            {
                TxtTotalWorkHours.Text = "Total Work Hours: 0:00";
                TxtAvgWorkHours.Text = "Average Work Hours: 0:00";
            }
        }

        private TimeSpan ParseWorkHours(string hoursStr)
        {
            if (string.IsNullOrEmpty(hoursStr)) return TimeSpan.Zero;

            try
            {
                var parts = hoursStr.Split(':');
                if (parts.Length == 2 && int.TryParse(parts[0], out int hours) && int.TryParse(parts[1], out int minutes))
                {
                    return new TimeSpan(hours, minutes, 0);
                }
            }
            catch { }

            return TimeSpan.Zero;
        }

        private string FormatMinutes(int totalMinutes)
        {
            var hours = totalMinutes / 60;
            var minutes = totalMinutes % 60;
            return $"{hours}:{minutes:D2}";
        }

        private void BtnRefreshMonth_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement monthly data refresh
            MessageBox.Show("Monthly data refresh functionality will be implemented soon!", "Info", 
                           MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnExportMonthly_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    DefaultExt = "csv",
                    FileName = $"MonthlyAttendance_{CurrentEmployeeId}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
                };

                if (saveDialog.ShowDialog() == true)
                {
                    var csv = new StringBuilder();
                    csv.AppendLine("Employee ID,Employee Name,Date,Swipe Count,In Time,Out Time,Total Hours,Work Hours,Status");
                    
                    foreach (var entry in MonthlyEntries)
                    {
                        csv.AppendLine($"\"{entry.EmployeeId}\",\"{entry.EmployeeName}\",\"{entry.Date}\",\"{entry.SwipeCount}\",\"{entry.InTime}\",\"{entry.OutTime}\",\"{entry.TotalHours}\",\"{entry.ActualWorkHours}\",\"{entry.Status}\"");
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

        private void BtnCopyMonthly_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var text = new StringBuilder();
                text.AppendLine("Employee ID\tEmployee Name\tDate\tSwipe Count\tIn Time\tOut Time\tTotal Hours\tWork Hours\tStatus");
                
                foreach (var entry in MonthlyEntries)
                {
                    text.AppendLine($"{entry.EmployeeId}\t{entry.EmployeeName}\t{entry.Date}\t{entry.SwipeCount}\t{entry.InTime}\t{entry.OutTime}\t{entry.TotalHours}\t{entry.ActualWorkHours}\t{entry.Status}");
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