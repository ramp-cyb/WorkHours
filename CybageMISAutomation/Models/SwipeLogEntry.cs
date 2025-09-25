namespace CybageMISAutomation.Models
{
    public class SwipeLogEntry
    {
        public string EmployeeId { get; set; } = string.Empty;
        public string EmployeeName { get; set; } = string.Empty;
        public string Date { get; set; } = string.Empty;
        public string InTime { get; set; } = string.Empty;
        public string OutTime { get; set; } = string.Empty;
        public string Duration { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
    }
}