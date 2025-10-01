namespace CybageMISAutomation.Models
{
    public class WorkHoursCalculation
    {
        public DateTime Date { get; set; }
        public string WorkingHoursDisplay { get; set; } = string.Empty;
        public string TotalWorkingHoursDisplay { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
    }
}
