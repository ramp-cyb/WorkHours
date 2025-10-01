using System.Collections.ObjectModel;

namespace CybageMISAutomation.Models
{
    public class FullReportViewModel
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public ObservableCollection<WeekRow> Weeks { get; set; } = new();
        public double TotalActualHours { get; set; }
        public double AverageActualHours { get; set; }
        public int WorkedDays { get; set; }
        public int LeaveDays { get; set; }
        public int WeeklyOffDays { get; set; }
        public int HolidayDays { get; set; }
        public int MissingDays { get; set; }
    }

    public class WeekRow
    {
        public List<DayCalendarItem> Days { get; set; } = new();
    }
}
