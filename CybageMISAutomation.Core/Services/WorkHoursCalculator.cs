using System.Globalization;
using CybageMISAutomation.Models;

namespace CybageMISAutomation.Services
{
    /// <summary>
    /// Labels swipe-log entries as WORK/PLAY and sums WORK time.
    /// Shared by MainWindow (calendar) and ComparisonWindow so both report the same hours.
    ///
    /// Rules:
    /// - Each swipe is placed on its own report Date + Time (never "today" by default).
    /// - Time after a swipe counts as WORK only when the swipe is at a work-floor reader
    ///   and is not an Exit. Time after any Exit, or after basement/parking/lobby readers, is PLAY.
    /// - "Ongoing" time (last swipe -> now) is added only when the last swipe is a WORK swipe
    ///   on the same calendar day as <c>now</c>, i.e. only for today's report.
    /// - Rows whose date/time cannot be parsed are skipped and reported, never treated as midnight.
    /// </summary>
    public static class WorkHoursCalculator
    {
        public const string Work = "WORK";
        public const string Play = "PLAY";

        private static readonly string[] DateFormats =
        {
            "d MMM yyyy", "dd MMM yyyy", "d-MMM-yyyy", "dd-MMM-yyyy", "d-MMM-yy", "dd-MMM-yy",
            "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "yyyy-MM-dd"
        };

        private static readonly string[] TimeFormats =
        {
            "h:mm:ss tt", "hh:mm:ss tt", "h:mm tt", "hh:mm tt", "H:mm:ss", "HH:mm:ss", "H:mm", "HH:mm"
        };

        // Same gate classification the app always used: basement/parking readers are outside the
        // work area, any other reader whose name contains "Floor" is a work-floor reader.
        private static readonly string[] OutsidePatterns = { "Basement", "Parking" };

        public sealed class Result
        {
            public List<SwipeLogEntry> Labeled { get; } = new();
            public List<string> SkippedRows { get; } = new();
            public TimeSpan WorkTime { get; set; }
            public TimeSpan OngoingTime { get; set; }
            public string WorkTimeDisplay => Format(WorkTime);
        }

        public static DateTime? ParseDate(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            return DateTime.TryParseExact(raw.Trim(), DateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces, out var d) ? d.Date : null;
        }

        public static TimeSpan? ParseTimeOfDay(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            return DateTime.TryParseExact(raw.Trim(), TimeFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault, out var t) ? t.TimeOfDay : null;
        }

        public static DateTime? ParseSwipe(SwipeLogEntry e)
        {
            var d = ParseDate(e.Date);
            var t = ParseTimeOfDay(e.Time);
            return d.HasValue && t.HasValue ? d.Value + t.Value : null;
        }

        public static bool IsExit(string? direction)
        {
            var d = (direction ?? string.Empty).Trim();
            return d.StartsWith("Exit", StringComparison.OrdinalIgnoreCase) ||
                   d.Equals("Out", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsWorkGate(string? machineName)
        {
            var m = machineName ?? string.Empty;
            if (OutsidePatterns.Any(p => m.Contains(p, StringComparison.OrdinalIgnoreCase)))
                return false;
            return m.Contains("Floor", StringComparison.OrdinalIgnoreCase);
        }

        public static string ClassifyActivity(SwipeLogEntry e) =>
            IsWorkGate(e.MachineName) && !IsExit(e.Direction) ? Work : Play;

        /// <param name="now">Current time; used only for ongoing time on the same day.</param>
        public static Result Calculate(IEnumerable<SwipeLogEntry>? entries, DateTime now)
        {
            var result = new Result();
            if (entries == null) return result;

            var parsed = new List<(SwipeLogEntry Entry, DateTime At)>();
            foreach (var e in entries)
            {
                var at = ParseSwipe(e);
                if (at.HasValue) parsed.Add((e, at.Value));
                else result.SkippedRows.Add($"{e.Date} {e.Time} {e.MachineName} {e.Direction}".Trim());
            }

            parsed.Sort((a, b) => a.At.CompareTo(b.At));

            var total = TimeSpan.Zero;
            for (int i = 0; i < parsed.Count; i++)
            {
                var (e, at) = parsed[i];
                var activity = ClassifyActivity(e);
                string duration;
                if (i < parsed.Count - 1)
                {
                    var chunk = parsed[i + 1].At - at;
                    duration = Format(chunk);
                    if (activity == Work) total += chunk;
                }
                else
                {
                    duration = "-";
                    if (activity == Work && at.Date == now.Date && now > at)
                    {
                        result.OngoingTime = now - at;
                        total += result.OngoingTime;
                    }
                }

                result.Labeled.Add(new SwipeLogEntry
                {
                    EmployeeId = e.EmployeeId,
                    Date = e.Date,
                    MachineName = e.MachineName,
                    Direction = e.Direction,
                    Time = e.Time,
                    Activity = activity,
                    Duration = duration
                });
            }

            result.WorkTime = total;
            return result;
        }

        public static string Format(TimeSpan span)
        {
            if (span < TimeSpan.Zero) span = TimeSpan.Zero;
            return $"{(int)span.TotalHours}:{span.Minutes:D2}";
        }
    }
}
