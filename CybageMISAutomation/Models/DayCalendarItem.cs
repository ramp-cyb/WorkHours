namespace CybageMISAutomation.Models
{
    public class DayCalendarItem
    {
        public DateTime Date { get; set; }
        public string Hours { get; set; } = string.Empty; // Display hours HH:MM
        public double HoursDecimal { get; set; } // For aggregate calculations
        public string Status { get; set; } = string.Empty;
        public bool FromMonthly { get; set; }
        public bool FromDailyFallback { get; set; }
        public int SwipeCount { get; set; }
        public string Tooltip { get; set; } = string.Empty;
        public bool IsToday { get; set; }
        public bool IsPlaceholder { get; set; } // For leading/trailing empty calendar cells
    }
}
