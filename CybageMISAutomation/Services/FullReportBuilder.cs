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
            if (!monthList.Any()) return new FullReportViewModel();

            // Determine month/year from first monthly entry if not provided
            DateTime firstDate = monthList
                .Select(m => ParseDate(m.Date))
                .Where(d => d != DateTime.MinValue)
                .OrderBy(d => d)
                .FirstOrDefault();
            int year = targetYear ?? firstDate.Year;
            int month = targetMonth ?? firstDate.Month;

            var dict = new Dictionary<DateTime, DayCalendarItem>();

            // 1. Seed with monthly
            foreach (var m in monthList)
            {
                var d = ParseDate(m.Date);
                if (d.Year != year || d.Month != month) continue;
                var hours = m.ActualWorkHours; // combined already chosen earlier
                var hoursDec = ParseHoursToDecimal(hours);
                var status = !string.IsNullOrWhiteSpace(m.Status) ? m.Status.Trim() : string.Empty;
                // Weekly off or leave derivation fallback
                if (string.IsNullOrEmpty(status) && hoursDec <= 0) status = "No Data";
                dict[d] = new DayCalendarItem
                {
                    Date = d,
                    Hours = hours,
                    HoursDecimal = hoursDec,
                    Status = status,
                    FromMonthly = true,
                    SwipeCount = m.SwipeCount,
                    Tooltip = BuildTooltip(hours, status, m)
                };
            }

            // Helper to inject from calculations (fallback + override blank monthly hours)
            void InjectDaily(IEnumerable<WorkHoursCalculation> src, bool isTodaySet)
            {
                if (src == null) return;
                foreach (var c in src)
                {
                    var d = c.Date.Date;
                    if (d.Year != year || d.Month != month) continue;
                    var hours = c.WorkingHoursDisplay ?? c.TotalWorkingHoursDisplay ?? string.Empty;
                    var hoursDec = ParseHoursToDecimal(hours);
                    if (dict.TryGetValue(d, out var existing))
                    {
                        // Always override today/yesterday with daily hours if we have them
                        if ((d.Date == DateTime.Today || d.Date == DateTime.Today.AddDays(-1)) && hoursDec > 0)
                        {
                            existing.Hours = hours;
                            existing.HoursDecimal = hoursDec;
                            existing.Status = "Present";
                            existing.FromDailyFallback = true;
                            existing.Tooltip = existing.Tooltip + "\n(Updated with Daily Hours: " + hours + ")";
                        }
                        // For other dates, only override if monthly has no hours
                        else if (existing.HoursDecimal <= 0 && hoursDec > 0)
                        {
                            existing.Hours = hours;
                            existing.HoursDecimal = hoursDec;
                            existing.Status = hoursDec > 0 ? "Present" : existing.Status;
                            existing.FromDailyFallback = true;
                            existing.Tooltip = existing.Tooltip + "\n(Overlaid Daily Hours: " + hours + ")";
                        }
                        continue; // do not add new item
                    }
                    var status = hoursDec > 0 ? "Present" : (IsWeekend(d) ? "Weekly Off" : "No Data");
                    dict[d] = new DayCalendarItem
                    {
                        Date = d,
                        Hours = hours,
                        HoursDecimal = hoursDec,
                        Status = status,
                        FromDailyFallback = true,
                        Tooltip = $"Source: Daily\nHours: {hours}\nStatus: {status}"
                    };
                }
            }

            InjectDaily(todayEntries, true);
            InjectDaily(yesterdayEntries, false);

            // Build calendar weeks
            var vm = new FullReportViewModel { Year = year, Month = month };
            var firstOfMonth = new DateTime(year, month, 1);
            int daysInMonth = DateTime.DaysInMonth(year, month);
            int leadingBlanks = ((int)firstOfMonth.DayOfWeek + 6) % 7; // Make Monday=0 if needed; using Monday start

            var allDays = new List<DayCalendarItem>();
            for (int i = 0; i < leadingBlanks; i++)
            {
                allDays.Add(new DayCalendarItem { IsPlaceholder = true });
            }
            for (int day = 1; day <= daysInMonth; day++)
            {
                var date = new DateTime(year, month, day);
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
            }

            // Pad trailing blanks to complete last week
            while (allDays.Count % 7 != 0) allDays.Add(new DayCalendarItem { IsPlaceholder = true });

            for (int i = 0; i < allDays.Count; i += 7)
            {
                vm.Weeks.Add(new WeekRow { Days = allDays.GetRange(i, 7) });
            }

            // Aggregations
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
            DateTime.TryParse(raw, out dt);
            return dt;
        }

        private static double ParseHoursToDecimal(string h)
        {
            if (string.IsNullOrWhiteSpace(h)) return 0;
            var parts = h.Split(':');
            if (parts.Length != 2) return 0;
            if (int.TryParse(parts[0], out int hh) && int.TryParse(parts[1], out int mm))
            {
                return hh + (mm / 60.0);
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
