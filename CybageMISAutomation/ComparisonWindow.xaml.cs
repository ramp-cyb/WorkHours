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
            // Process and label Yesterday data
            var labeledYesterdayData = LabelSwipeEntries(yesterdayData);
            YesterdayEntries.Clear();
            foreach (var entry in labeledYesterdayData)
            {
                YesterdayEntries.Add(entry);
            }
            
            // Process and label Today data  
            var labeledTodayData = LabelSwipeEntries(todayData);
            TodayEntries.Clear();
            foreach (var entry in labeledTodayData)
            {
                TodayEntries.Add(entry);
            }
            
            // Update counts and calculate work hours
            TxtYesterdayCount.Text = $"Yesterday: {yesterdayData.Count} records";
            TxtTodayCount.Text = $"Today: {todayData.Count} records";
            
            // Calculate work hours based on labeled time chunks
            var yesterdayWorkHours = CalculateWorkHoursFromLabels(labeledYesterdayData);
            var todayWorkHours = CalculateWorkHoursFromLabels(labeledTodayData);
            
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

        private string IdentifyGateType(string machineName)
        {
            if (string.IsNullOrEmpty(machineName))
                return "Unknown";

            // Based on CT2 building structure and VBA gate classification
            // Main gates - building/campus entry points (like Basement, Parking)
            var mainGatePatterns = new[] { "Basement", "Parking", "Tripod", "Ground", "Reception", "Main Gate", "Security", "Entry Gate" };
            
            // Work gates - actual office/work floors (4th Floor, 5th Floor, etc.)
            var workGatePatterns = new[] { "4th Floor", "5th Floor", "6th Floor", "7th Floor", "8th Floor", "9th Floor", "Floor Le", "Floor Ri", "Building.*Floor", "Office Floor" };
            
            // Play gates - recreation areas (cafeteria, lobby, etc.)
            var playGatePatterns = new[] { "Cafeteria", "Recreation", "Lobby", "Canteen", "Food Court", "Gaming", "Break" };

            // Check in order of specificity
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

        private List<SwipeLogEntry> LabelSwipeEntries(List<SwipeLogEntry> entries)
        {
            if (entries == null || entries.Count == 0)
                return entries ?? new List<SwipeLogEntry>();

            var labeledEntries = new List<SwipeLogEntry>();
            
            // Sort entries by time 
            var sortedEntries = entries.OrderBy(e => ParseTime(e.Time)).ToList();
            
            System.Diagnostics.Debug.WriteLine($"\\n=== LABELING {entries.Count} SWIPE ENTRIES AND CALCULATING DURATIONS ===");
            
            for (int i = 0; i < sortedEntries.Count; i++)
            {
                var entry = sortedEntries[i];
                var gateType = IdentifyGateType(entry.MachineName);
                var direction = entry.Direction;
                var activity = "";
                var duration = "";
                
                // Apply VBA labeling logic - based on your clarification:
                // WORK = inside work area (office floors)
                // PLAY = outside work area (basement, parking, reception, etc.)
                if (gateType == "MainGate" && (entry.MachineName.Contains("Basement") || entry.MachineName.Contains("Parking")))
                {
                    // Real exit gates - basement/parking areas are PLAY (outside work area)
                    activity = "PLAY"; // Outside work area
                }
                else if (gateType == "WorkGate" || entry.MachineName.Contains("Floor"))
                {
                    // Floor gates - WORK area (inside work area)
                    activity = "WORK";
                }
                else if (gateType == "PlayGate")
                {
                    // Recreation areas - PLAY (outside work area)
                    activity = "PLAY";
                }
                else if (gateType == "MainGate")
                {
                    // Other main gates (reception, lobby) - PLAY (outside work area)
                    activity = "PLAY";
                }
                else
                {
                    // Unknown areas - default to PLAY (outside work area)
                    activity = "PLAY";
                }
                
                // Calculate duration from this swipe to next swipe
                if (i < sortedEntries.Count - 1)
                {
                    var currentTime = ParseTime(entry.Time);
                    var nextTime = ParseTime(sortedEntries[i + 1].Time);
                    
                    if (nextTime > currentTime)
                    {
                        var timeDiff = nextTime - currentTime;
                        var totalMinutes = (int)timeDiff.TotalMinutes;
                        var hours = totalMinutes / 60;
                        var minutes = totalMinutes % 60;
                        duration = $"{hours}:{minutes:D2}";
                    }
                    else
                    {
                        duration = "0:00";
                    }
                }
                else
                {
                    duration = "-"; // Last entry has no next swipe
                }
                
                // Create labeled entry with duration
                var labeledEntry = new SwipeLogEntry
                {
                    EmployeeId = entry.EmployeeId,
                    Date = entry.Date,
                    MachineName = entry.MachineName,
                    Direction = entry.Direction,
                    Time = entry.Time,
                    Activity = activity,
                    Duration = duration
                };
                
                System.Diagnostics.Debug.WriteLine($"{entry.Time} | {direction} | {entry.MachineName} | {gateType} → Activity: {activity} | Duration: {duration}");
                labeledEntries.Add(labeledEntry);
                
                System.Diagnostics.Debug.WriteLine($"{entry.Time} | {direction} | {entry.MachineName} | {gateType} → '{activity}'");
            }
            
            return labeledEntries;
        }

        private string CalculateWorkHoursFromLabels(List<SwipeLogEntry> labeledEntries)
        {
            if (labeledEntries == null || labeledEntries.Count == 0)
                return "0:00";

            try
            {
                var totalWorkMinutes = 0.0;
                
                System.Diagnostics.Debug.WriteLine($"\\n=== CALCULATING WORK TIME FROM LABELS ===");
                
                // Calculate time chunks between consecutive swipes
                for (int i = 0; i < labeledEntries.Count - 1; i++)
                {
                    var currentEntry = labeledEntries[i];
                    var nextEntry = labeledEntries[i + 1];
                    
                    // The time chunk between current and next swipe has the activity of current EXIT
                    if (currentEntry.Activity == "WORK")
                    {
                        var startTime = ParseTime(currentEntry.Time);
                        var endTime = ParseTime(nextEntry.Time);
                        
                        if (endTime > startTime)
                        {
                            var chunkMinutes = (endTime - startTime).TotalMinutes;
                            totalWorkMinutes += chunkMinutes;
                            
                            System.Diagnostics.Debug.WriteLine($"WORK chunk: {startTime:HH:mm:ss} to {endTime:HH:mm:ss} = {chunkMinutes:F0} minutes");
                        }
                    }
                }
                
                // Handle ongoing work for today (if last entry indicates work)
                if (labeledEntries.Count > 0)
                {
                    var lastEntry = labeledEntries.Last();
                    if (lastEntry.Activity == "WORK")
                    {
                        var lastTime = ParseTime(lastEntry.Time);
                        var now = DateTime.Now;
                        
                        // Only add ongoing time if it's the same day
                        if (now.Date == lastTime.Date && now > lastTime)
                        {
                            var ongoingMinutes = (now - lastTime).TotalMinutes;
                            totalWorkMinutes += ongoingMinutes;
                            
                            System.Diagnostics.Debug.WriteLine($"ONGOING work: {lastTime:HH:mm:ss} to {now:HH:mm:ss} = {ongoingMinutes:F0} minutes");
                        }
                    }
                }
                
                var hours = (int)(totalWorkMinutes / 60);
                var minutes = (int)(totalWorkMinutes % 60);
                
                System.Diagnostics.Debug.WriteLine($"TOTAL WORK TIME: {totalWorkMinutes:F0} minutes = {hours}:{minutes:D2}\\n");
                
                return $"{hours}:{minutes:D2}";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error calculating work hours: {ex.Message}");
                return "Error";
            }
        }
    }
}