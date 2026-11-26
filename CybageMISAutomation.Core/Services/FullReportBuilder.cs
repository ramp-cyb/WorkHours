using CybageMISAutomation.Models;

namespace CybageMISAutomation.Services
{
    public static class FullReportBuilder
    {
        public static FullReportViewModel Build(
            IEnumerable<MonthlyAttendanceEntry> monthly,
            IEnumerable<WorkHoursCalculation> todayEntries,
            IEnumerable<WorkHoursCalculation> yesterdayEntries,
            int? targetYear = null,
            int? targetMonth = null)
        {
            var monthList = monthly?.ToList() ?? new List<MonthlyAttendanceEntry>();

            // The calendar is for the current month unless told otherwise; it is built even when
            // the monthly report is missing so Today/Yesterday results are never thrown away.
            int year = targetYear ?? DateTime.Today.Year;
            int month = targetMonth ?? DateTime.Today.Month;

            var dict = new Dictionary<DateTime, DayCalendarItem>();

            foreach (var m in monthList)
            {
                var d = ParseDate(m.Date);
                if (d == DateTime.MinValue || d.Year != year || d.Month != month) continue;
                var hours = m.ActualWorkHours;
                var hoursDec = ParseHoursToDecimal(hours);
                var status = !string.IsNullOrWhiteSpace(m.Status) ? m.Status.Trim() : string.Empty;
                if (string.IsNullOrEmpty(status) && hoursDec <= 0) status = "No Data";

                // Last row for a date wins (deterministic).
                dict[d.Date] = new DayCalendarItem
                {
                    Date = d.Date,
                    // Show every day as H:mm (monthly MIS gives decimal hours); the raw value stays in the tooltip.
                    Hours = hoursDec > 0 ? WorkHoursCalculator.Format(TimeSpan.FromHours(hoursDec)) : hours,
                    HoursDecimal = hoursDec,
                    Status = status,
                    FromMonthly = true,
                    SwipeCount = m.SwipeCount,
                    Tooltip = BuildTooltip(hours, status, m)
                };
            }

            void InjectDaily(IEnumerable<WorkHoursCalculation> src)
            {
                if (src == null) return;
                foreach (var c in src)
                {
                    var d = c.Date.Date;
                    if (d.Year != year || d.Month != month) continue;
                    var hours = !string.IsNullOrWhiteSpace(c.WorkingHoursDisplay) ? c.WorkingHoursDisplay : c.TotalWorkingHoursDisplay ?? string.Empty;
                    var hoursDec = ParseHoursToDecimal(hours);
                    if (hoursDec <= 0) continue; // no swipe-based work time: keep whatever monthly says

                    if (dict.TryGetValue(d, out var existing))
                    {
                        // Today is always in progress in the monthly report, so the live swipe value wins.
                        // For other days the monthly figure (which also includes WFH) is authoritative
                        // and is only filled in from the swipe log when it has no hours yet.
                        if (d == DateTime.Today || existing.HoursDecimal <= 0)
                        {
                            existing.Hours = hours;
                            existing.HoursDecimal = hoursDec;
                            existing.Status = "Present";
                            existing.FromDailyFallback = true;
                            existing.Tooltip += $"\n(Swipe-log hours: {hours})";
                        }
                    }
                    else
                    {
                        dict[d] = new DayCalendarItem
                        {
                            Date = d,
                            Hours = hours,
                            HoursDecimal = hoursDec,
                            Status = "Present",
                            FromDailyFallback = true,
                            Tooltip = $"Source: Swipe log\nHours: {hours}"
                        };
                    }
                }
            }

            InjectDaily(todayEntries);
            InjectDaily(yesterdayEntries);

            var vm = new FullReportViewModel { Year = year, Month = month };
            var firstOfMonth = new DateTime(year, month, 1);
            var lastOfMonth = firstOfMonth.AddMonths(1).AddDays(-1);
            var finalDate = dict.Keys.Any() ? dict.Keys.Max() : lastOfMonth;
            if (finalDate < lastOfMonth)
            {
                finalDate = lastOfMonth;
            }
            int leadingBlanks = ((int)firstOfMonth.DayOfWeek + 6) % 7;

            var allDays = new List<DayCalendarItem>();
            for (int i = 0; i < leadingBlanks; i++)
            {
                allDays.Add(new DayCalendarItem { IsPlaceholder = true });
            }
            var cursor = firstOfMonth;
            while (cursor <= finalDate)
            {
                var date = cursor;
                if (!dict.TryGetValue(date, out var item))
                {
                    item = new DayCalendarItem
                    {
                        Date = date,
                        Hours = string.Empty,
                        HoursDecimal = 0,
                        Status = IsWeekend(date) ? "Weekly Off" : "No Data",
                        Tooltip = IsWeekend(date) ? "Weekend" : "No data",
                        IsToday = date.Date == DateTime.Today
                    };
                }
                else
                {
                    item.IsToday = date.Date == DateTime.Today;
                    if (string.IsNullOrWhiteSpace(item.Status))
                    {
                        item.Status = item.HoursDecimal > 0 ? "Present" : (IsWeekend(date) ? "Weekly Off" : "No Data");
                    }
                }
                allDays.Add(item);
                cursor = cursor.AddDays(1);
            }

            while (allDays.Count % 7 != 0) allDays.Add(new DayCalendarItem { IsPlaceholder = true });

            for (int i = 0; i < allDays.Count; i += 7)
            {
                vm.Weeks.Add(new WeekRow { Days = allDays.GetRange(i, 7) });
            }

            var realDays = allDays.Where(d => !d.IsPlaceholder && d.Date != DateTime.MinValue);
            vm.TotalActualHours = realDays.Sum(d => d.HoursDecimal);
            var worked = realDays.Where(d => d.Status.Equals("Present", StringComparison.OrdinalIgnoreCase));
            vm.WorkedDays = worked.Count();
            vm.LeaveDays = realDays.Count(d => d.Status.Contains("Leave", StringComparison.OrdinalIgnoreCase));
            vm.WeeklyOffDays = realDays.Count(d => d.Status.Contains("Weekly Off", StringComparison.OrdinalIgnoreCase));
            vm.HolidayDays = realDays.Count(d => d.Status.Contains("Holiday", StringComparison.OrdinalIgnoreCase));
            int countedDays = realDays.Count(d => d.HoursDecimal > 0);
            vm.AverageActualHours = countedDays > 0 ? vm.TotalActualHours / countedDays : 0;
            vm.MissingDays = realDays.Count(d => d.Status == "No Data");

            return vm;
        }

        private static DateTime ParseDate(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return DateTime.MinValue;
            string[] formats = { "dd-MMM-yyyy", "d-MMM-yyyy", "dd-MMM-yy", "d-MMM-yy", "dd-MM-yyyy", "d-MM-yyyy" };
            if (DateTime.TryParseExact(raw.Trim(), formats, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out var dt)) return dt;
            // Same formats the swipe log uses ("25 Sep 2025" etc.); never guess with the machine's culture.
            return WorkHoursCalculator.ParseDate(raw) ?? DateTime.MinValue;
        }

        /// <summary>Parses "H:mm", "H:mm:ss" or decimal hours ("8.49") to decimal hours; 0 when unreadable.</summary>
        public static double ParseHoursToDecimal(string? h)
        {
            if (string.IsNullOrWhiteSpace(h)) return 0;
            h = h.Trim();

            // "H:mm" or "H:mm:ss" (daily swipe-log calculations)
            if (h.Contains(':'))
            {
                var parts = h.Split(':');
                if ((parts.Length == 2 || parts.Length == 3) &&
                    int.TryParse(parts[0], out int hh) && int.TryParse(parts[1], out int mm))
                {
                    int ss = parts.Length == 3 && int.TryParse(parts[2], out int s) ? s : 0;
                    return hh + (mm / 60.0) + (ss / 3600.0);
                }
                return 0;
            }

            // Decimal hours, e.g. "8.49" (monthly MIS report; values like "7.68" confirm it's not h.mm)
            if (double.TryParse(h, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out double dec) && dec > 0)
            {
                return dec;
            }
            return 0;
        }

        private static bool IsWeekend(DateTime dt) => dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday;

        private static string BuildTooltip(string hours, string status, MonthlyAttendanceEntry m)
        {
            return $"Source: Monthly\nHours: {hours}\nStatus: {status}\nSwipe Count: {m.SwipeCount}\nWFH Actual: {m.ActualWFHHours}\nSwipe Actual: {m.ActualWorkHours}\nTotal Swipe: {m.TotalHours}";
        }
    }
}
