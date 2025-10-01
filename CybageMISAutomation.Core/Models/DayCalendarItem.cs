namespace CybageMISAutomation.Models
{
    public class DayCalendarItem
    {
        public DateTime Date { get; set; }
        public string Hours { get; set; } = string.Empty;
        public double HoursDecimal { get; set; }
        public string Status { get; set; } = string.Empty;
        public bool FromMonthly { get; set; }
        public bool FromDailyFallback { get; set; }
        public int SwipeCount { get; set; }
        public string Tooltip { get; set; } = string.Empty;
        public bool IsToday { get; set; }
        public bool IsPlaceholder { get; set; }
    }
}
