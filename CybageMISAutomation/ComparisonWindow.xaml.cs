using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using CybageMISAutomation.Models;
using System.Globalization;

namespace CybageMISAutomation
{
    public partial class ComparisonWindow : Window
    {
        public ObservableCollection<SwipeLogEntry> YesterdayEntries { get; set; }
        public ObservableCollection<SwipeLogEntry> TodayEntries { get; set; }
        private MainWindow? _mainWindow;
        private WorkHoursCalculation _workHoursCalculator = new WorkHoursCalculation();

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
            
            // Calculate and display work hours
            CalculateAndDisplayWorkHours(yesterdayData, todayData);
        }
        
        private void CalculateAndDisplayWorkHours(List<SwipeLogEntry> yesterdayData, List<SwipeLogEntry> todayData)
        {
            // Calculate Yesterday work hours
            var yesterdaySwipeEntries = ConvertToSwipeEntries(yesterdayData);
            var yesterdaySessions = _workHoursCalculator.CalculateWorkSessions(yesterdaySwipeEntries);
            var yesterdayWorkHours = _workHoursCalculator.GetTotalWorkHours(yesterdaySessions);
            var yesterdayPlayHours = _workHoursCalculator.GetTotalPlayHours(yesterdaySessions);
            
            // Calculate Today work hours
            var todaySwipeEntries = ConvertToSwipeEntries(todayData);
            var todaySessions = _workHoursCalculator.CalculateWorkSessions(todaySwipeEntries);
            var todayWorkHours = _workHoursCalculator.GetTotalWorkHours(todaySessions);
            var todayPlayHours = _workHoursCalculator.GetTotalPlayHours(todaySessions);
            
            // Update UI
            TxtYesterdayWorkHours.Text = _workHoursCalculator.FormatDuration(yesterdayWorkHours);
            TxtYesterdayPlayHours.Text = _workHoursCalculator.FormatDuration(yesterdayPlayHours);
            TxtTodayWorkHours.Text = _workHoursCalculator.FormatDuration(todayWorkHours);
            TxtTodayPlayHours.Text = _workHoursCalculator.FormatDuration(todayPlayHours);
        }
        
        private List<SwipeEntry> ConvertToSwipeEntries(List<SwipeLogEntry> swipeLogData)
        {
            var swipeEntries = new List<SwipeEntry>();
            
            foreach (var entry in swipeLogData)
            {
                if (!string.IsNullOrEmpty(entry.SwipeTime) && !string.IsNullOrEmpty(entry.Date))
                {
                    if (DateTime.TryParse($"{entry.Date} {entry.SwipeTime}", out DateTime swipeDateTime))
                    {
                        swipeEntries.Add(new SwipeEntry
                        {
                            EmployeeId = entry.EmployeeId,
                            SwipeDateTime = swipeDateTime,
                            Gate = entry.Gate,
                            Direction = entry.Direction,
                            TimeString = entry.SwipeTime,
                            DateString = entry.Date
                        });
                    }
                }
            }
            
            return swipeEntries;
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
                    csv.AppendLine("Report Type,Employee ID,Date,Gate,Direction,Swipe Time,Work Hours,Play Hours");
                    
                    // Yesterday data
                    foreach (var entry in YesterdayEntries)
                    {
                        csv.AppendLine($"Yesterday,\"{entry.EmployeeId}\",\"{entry.Date}\",\"{entry.Gate}\",\"{entry.Direction}\",\"{entry.SwipeTime}\",\"{entry.CalculatedWorkHours}\",\"{entry.CalculatedPlayHours}\"");
                    }
                    
                    // Today data
                    foreach (var entry in TodayEntries)
                    {
                        csv.AppendLine($"Today,\"{entry.EmployeeId}\",\"{entry.Date}\",\"{entry.Gate}\",\"{entry.Direction}\",\"{entry.SwipeTime}\",\"{entry.CalculatedWorkHours}\",\"{entry.CalculatedPlayHours}\"");
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
                text.AppendLine("Report Type\tEmployee ID\tDate\tGate\tDirection\tSwipe Time\tWork Hours\tPlay Hours");
                
                // Yesterday data
                foreach (var entry in YesterdayEntries)
                {
                    text.AppendLine($"Yesterday\t{entry.EmployeeId}\t{entry.Date}\t{entry.Gate}\t{entry.Direction}\t{entry.SwipeTime}\t{entry.CalculatedWorkHours}\t{entry.CalculatedPlayHours}");
                }
                
                // Today data
                foreach (var entry in TodayEntries)
                {
                    text.AppendLine($"Today\t{entry.EmployeeId}\t{entry.Date}\t{entry.Gate}\t{entry.Direction}\t{entry.SwipeTime}\t{entry.CalculatedWorkHours}\t{entry.CalculatedPlayHours}");
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