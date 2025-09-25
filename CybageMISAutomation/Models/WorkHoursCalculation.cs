using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CybageMISAutomation.Models
{
    public class SwipeEntry
    {
        public string EmployeeId { get; set; } = string.Empty;
        public DateTime SwipeDateTime { get; set; }
        public string Gate { get; set; } = string.Empty;
        public string Direction { get; set; } = string.Empty; // Entry or Exit
        public GateType GateType { get; set; }
        public string TimeString { get; set; } = string.Empty;
        public string DateString { get; set; } = string.Empty;
    }

    public enum GateType
    {
        MainGate,
        PlayGate,
        WorkGate,
        Unknown
    }

    public enum AreaType
    {
        Work,
        Play,
        Campus,
        Outside
    }

    public class WorkSession
    {
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public AreaType Area { get; set; }
        public TimeSpan Duration => EndTime - StartTime;
    }

    public class WorkHoursCalculation
    {
        private readonly string[] _mainGates;
        private readonly string[] _playGates;
        private readonly string[] _workGates;

        public WorkHoursCalculation()
        {
            // Based on VBA configuration - these would ideally be configurable
            _mainGates = new[] { "Main Gate", "Campus Gate", "Security Gate" };
            _playGates = new[] { "Play Gate", "Recreation Gate", "Cafeteria Gate" };
            _workGates = new[] { "Work Gate", "Office Gate", "Floor Gate" };
        }

        public WorkHoursCalculation(string[] mainGates, string[] playGates, string[] workGates)
        {
            _mainGates = mainGates ?? new string[0];
            _playGates = playGates ?? new string[0];
            _workGates = workGates ?? new string[0];
        }

        public GateType IdentifyGateType(string gateName)
        {
            if (string.IsNullOrEmpty(gateName))
                return GateType.Unknown;

            if (_mainGates.Any(gate => gateName.Contains(gate, StringComparison.OrdinalIgnoreCase)))
                return GateType.MainGate;

            if (_playGates.Any(gate => gateName.Contains(gate, StringComparison.OrdinalIgnoreCase)))
                return GateType.PlayGate;

            if (_workGates.Any(gate => gateName.Contains(gate, StringComparison.OrdinalIgnoreCase)))
                return GateType.WorkGate;

            return GateType.Unknown;
        }

        public List<WorkSession> CalculateWorkSessions(List<SwipeEntry> swipeEntries)
        {
            var sessions = new List<WorkSession>();
            if (!swipeEntries.Any()) return sessions;

            // Sort by time
            var sortedEntries = swipeEntries.OrderBy(e => e.SwipeDateTime).ToList();

            // Process each swipe entry to determine area transitions
            var processedEntries = new List<SwipeEntry>();
            
            foreach (var entry in sortedEntries)
            {
                entry.GateType = IdentifyGateType(entry.Gate);
                
                // Mark area based on direction and gate type (following VBA logic)
                MarkAreaTransition(entry);
                processedEntries.Add(entry);
            }

            // Calculate sessions based on area transitions
            AreaType currentArea = AreaType.Outside;
            DateTime? sessionStart = null;

            foreach (var entry in processedEntries)
            {
                var newArea = DetermineAreaAfterSwipe(entry, currentArea);
                
                if (newArea != currentArea)
                {
                    // End current session if there was one
                    if (sessionStart.HasValue && (currentArea == AreaType.Work || currentArea == AreaType.Play))
                    {
                        sessions.Add(new WorkSession
                        {
                            StartTime = sessionStart.Value,
                            EndTime = entry.SwipeDateTime,
                            Area = currentArea
                        });
                    }

                    // Start new session if entering work or play area
                    if (newArea == AreaType.Work || newArea == AreaType.Play)
                    {
                        sessionStart = entry.SwipeDateTime;
                    }
                    else
                    {
                        sessionStart = null;
                    }

                    currentArea = newArea;
                }
            }

            // If still in a session at the end, close it with current time (for today's data)
            if (sessionStart.HasValue && (currentArea == AreaType.Work || currentArea == AreaType.Play))
            {
                var lastEntry = processedEntries.LastOrDefault();
                var endTime = lastEntry?.SwipeDateTime.Date == DateTime.Today ? DateTime.Now : lastEntry?.SwipeDateTime ?? DateTime.Now;
                
                sessions.Add(new WorkSession
                {
                    StartTime = sessionStart.Value,
                    EndTime = endTime,
                    Area = currentArea
                });
            }

            return sessions;
        }

        private void MarkAreaTransition(SwipeEntry entry)
        {
            // This follows the VBA logic for marking areas based on direction and gate type
            if (entry.Direction == "Exit" && entry.GateType == GateType.MainGate)
            {
                // Leaving campus - no specific area marking needed
            }
            else if (entry.Direction == "Exit" && entry.GateType == GateType.PlayGate)
            {
                // Leaving play area - previous time was PLAY
            }
            else if (entry.Direction == "Exit" && entry.GateType == GateType.WorkGate)
            {
                // Leaving work area - previous time was WORK
            }
            else if (entry.Direction == "Entry" && entry.GateType == GateType.WorkGate)
            {
                // Entering work area - previous time was PLAY
            }
        }

        private AreaType DetermineAreaAfterSwipe(SwipeEntry entry, AreaType currentArea)
        {
            switch (entry.GateType)
            {
                case GateType.MainGate:
                    return entry.Direction == "Entry" ? AreaType.Campus : AreaType.Outside;
                
                case GateType.PlayGate:
                    return entry.Direction == "Entry" ? AreaType.Play : AreaType.Campus;
                
                case GateType.WorkGate:
                    return entry.Direction == "Entry" ? AreaType.Work : AreaType.Play;
                
                default:
                    return currentArea;
            }
        }

        public TimeSpan GetTotalWorkHours(List<WorkSession> sessions)
        {
            return TimeSpan.FromTicks(sessions.Where(s => s.Area == AreaType.Work).Sum(s => s.Duration.Ticks));
        }

        public TimeSpan GetTotalPlayHours(List<WorkSession> sessions)
        {
            return TimeSpan.FromTicks(sessions.Where(s => s.Area == AreaType.Play).Sum(s => s.Duration.Ticks));
        }

        public string FormatDuration(TimeSpan duration)
        {
            return $"{(int)duration.TotalHours:00}:{duration.Minutes:00}";
        }
    }

    public class DailyWorkSummary
    {
        public DateTime Date { get; set; }
        public TimeSpan TotalWorkHours { get; set; }
        public TimeSpan TotalPlayHours { get; set; }
        public TimeSpan TotalTimeInCampus { get; set; }
        public List<WorkSession> Sessions { get; set; } = new List<WorkSession>();
        public List<SwipeEntry> SwipeEntries { get; set; } = new List<SwipeEntry>();
        
        public string FormattedWorkHours => FormatTime(TotalWorkHours);
        public string FormattedPlayHours => FormatTime(TotalPlayHours);
        public string FormattedTotalHours => FormatTime(TotalTimeInCampus);
        
        private string FormatTime(TimeSpan time)
        {
            return $"{(int)time.TotalHours:00}:{time.Minutes:00}";
        }
    }
}