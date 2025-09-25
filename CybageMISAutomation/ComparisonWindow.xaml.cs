using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using Microsoft.Win32;
using CybageMISAutomation.Models;
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
            
            // Update counts and calculate work hours
            TxtYesterdayCount.Text = $"Yesterday: {yesterdayData.Count} records";
            TxtTodayCount.Text = $"Today: {todayData.Count} records";
            
            // Calculate work hours for both days
            var yesterdayWorkHours = CalculateWorkHours(yesterdayData);
            var todayWorkHours = CalculateWorkHours(todayData);
            
            TxtYesterdayWorkHours.Text = $"Yesterday Work Hours: {yesterdayWorkHours}";
            TxtTodayWorkHours.Text = $"Today Work Hours: {todayWorkHours}";
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
                    csv.AppendLine("Report Type,Employee ID,Date,Machine Name,Direction,Time");
                    
                    // Yesterday data
                    foreach (var entry in YesterdayEntries)
                    {
                        csv.AppendLine($"Yesterday,\"{entry.EmployeeId}\",\"{entry.Date}\",\"{entry.MachineName}\",\"{entry.Direction}\",\"{entry.Time}\"");
                    }
                    
                    // Today data
                    foreach (var entry in TodayEntries)
                    {
                        csv.AppendLine($"Today,\"{entry.EmployeeId}\",\"{entry.Date}\",\"{entry.MachineName}\",\"{entry.Direction}\",\"{entry.Time}\"");
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
                text.AppendLine("Report Type\tEmployee ID\tDate\tMachine Name\tDirection\tTime");
                
                // Yesterday data
                foreach (var entry in YesterdayEntries)
                {
                    text.AppendLine($"Yesterday\t{entry.EmployeeId}\t{entry.Date}\t{entry.MachineName}\t{entry.Direction}\t{entry.Time}");
                }
                
                // Today data
                foreach (var entry in TodayEntries)
                {
                    text.AppendLine($"Today\t{entry.EmployeeId}\t{entry.Date}\t{entry.MachineName}\t{entry.Direction}\t{entry.Time}");
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

        private string CalculateWorkHours(List<SwipeLogEntry> entries)
        {
            if (entries == null || entries.Count == 0)
                return "0:00";

            try
            {
                // Sort entries by time to process chronologically
                var sortedEntries = entries.OrderBy(e => ParseTime(e.Time)).ToList();
                
                // Debug output to understand the data
                System.Diagnostics.Debug.WriteLine($"\\n=== CALCULATING WORK HOURS FOR {entries.Count} ENTRIES ===");
                foreach (var entry in sortedEntries)
                {
                    var gateType = IdentifyGateType(entry.MachineName);
                    System.Diagnostics.Debug.WriteLine($"{ParseTime(entry.Time):HH:mm:ss} | {entry.Direction} | {entry.MachineName} | GateType: {gateType}");
                }
                
                var workPeriods = new List<(DateTime start, DateTime end)>();
                DateTime? workStartTime = null;
                
                foreach (var entry in sortedEntries)
                {
                    var gateType = IdentifyGateType(entry.MachineName);
                    var direction = entry.Direction;
                    
                    // Corrected logic for swipe card systems:
                    // Entry + WorkGate = Entering work floor (start work period)
                    // Exit + WorkGate = Leaving work floor (end work period)
                    // MainGate/PlayGate = Campus/building access (not counted as work time)
                    
                    if (direction == "Entry" && gateType == "WorkGate")
                    {
                        // Starting work period (entering work floor)
                        if (workStartTime == null)
                        {
                            workStartTime = ParseTime(entry.Time);
                        }
                    }
                    else if (direction == "Exit" && gateType == "WorkGate")
                    {
                        // Ending work period (leaving work floor)
                        if (workStartTime.HasValue)
                        {
                            var endTime = ParseTime(entry.Time);
                            if (endTime > workStartTime.Value)
                            {
                                workPeriods.Add((workStartTime.Value, endTime));
                            }
                            workStartTime = null;
                        }
                    }
                    // Ignore MainGate and PlayGate entries for work time calculation
                }
                
                // If still in work area at end of day, add current time as end
                if (workStartTime.HasValue)
                {
                    var now = DateTime.Now;
                    var todayDate = ParseTime(sortedEntries[0].Time).Date;
                    
                    // If it's the same day, use current time, otherwise use end of day
                    var endTime = (now.Date == todayDate) ? now : todayDate.AddHours(18); // Assume 6 PM end
                    workPeriods.Add((workStartTime.Value, endTime));
                }
                
                // Debug work periods
                System.Diagnostics.Debug.WriteLine($"\\nWork Periods Found: {workPeriods.Count}");
                foreach (var period in workPeriods)
                {
                    var duration = period.end - period.start;
                    System.Diagnostics.Debug.WriteLine($"  {period.start:HH:mm:ss} to {period.end:HH:mm:ss} = {duration.TotalHours:F2} hours");
                }
                
                // Calculate total work hours
                var totalMinutes = workPeriods.Sum(period => 
                    (period.end - period.start).TotalMinutes);
                
                var hours = (int)(totalMinutes / 60);
                var minutes = (int)(totalMinutes % 60);
                
                System.Diagnostics.Debug.WriteLine($"Total: {totalMinutes} minutes = {hours}:{minutes:D2}\\n");
                
                return $"{hours}:{minutes:D2}";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error calculating work hours: {ex.Message}");
                return "Error calculating";
            }
        }

        private string IdentifyGateType(string machineName)
        {
            if (string.IsNullOrEmpty(machineName))
                return "Unknown";

            // Based on actual CT2 building data patterns
            // Work gates (actual work floor access) - floors above ground
            var workGatePatterns = new[] { "Floor", "4th Floor", "5th Floor", "6th Floor", "7th Floor", "8th Floor", "9th Floor", "Building.*Floor", "Office", "Development", "Lab", "Studio" };
            
            // Main gates (campus/building entry/exit) - ground level access
            var mainGatePatterns = new[] { "Main Gate", "Security Gate", "Reception", "Parking", "Tripod", "Basement", "Ground", "Entry", "Exit Gate" };
            
            // Play gates (cafeteria, recreation areas)
            var playGatePatterns = new[] { "Cafeteria", "Recreation", "Lobby", "Canteen", "Food Court", "Gaming" };

            // Check work gates first (more specific)
            foreach (var pattern in workGatePatterns)
            {
                if (machineName.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                    return "WorkGate";
            }

            foreach (var pattern in playGatePatterns)
            {
                if (machineName.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                    return "PlayGate";
            }

            foreach (var pattern in mainGatePatterns)
            {
                if (machineName.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                    return "MainGate";
            }

            return "Unknown";
        }

        private DateTime ParseTime(string timeStr)
        {
            if (string.IsNullOrEmpty(timeStr))
                return DateTime.MinValue;

            // Try to parse various time formats
            var formats = new[]
            {
                "HH:mm:ss",
                "H:mm:ss", 
                "HH:mm",
                "H:mm",
                "dd/MM/yyyy HH:mm:ss",
                "dd-MM-yyyy HH:mm:ss",
                "MM/dd/yyyy HH:mm:ss",
                "yyyy-MM-dd HH:mm:ss"
            };

            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(timeStr, format, null, System.Globalization.DateTimeStyles.None, out DateTime result))
                {
                    // If only time was parsed, assume today's date
                    if (result.Date == DateTime.MinValue.Date)
                    {
                        result = DateTime.Today.Add(result.TimeOfDay);
                    }
                    return result;
                }
            }

            // Fallback to general parsing
            if (DateTime.TryParse(timeStr, out DateTime generalResult))
            {
                return generalResult;
            }

            return DateTime.MinValue;
        }
    }
}