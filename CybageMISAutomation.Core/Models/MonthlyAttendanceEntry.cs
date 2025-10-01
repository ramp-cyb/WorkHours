namespace CybageMISAutomation.Models
{
    public class MonthlyAttendanceEntry
    {
        public string EmployeeId { get; set; } = string.Empty;
        public string EmployeeName { get; set; } = string.Empty;
        public string Date { get; set; } = string.Empty;
        public int SwipeCount { get; set; }
        public string InTime { get; set; } = string.Empty;
        public string OutTime { get; set; } = string.Empty;
        public string TotalHours { get; set; } = string.Empty;
        public string ActualWorkHours { get; set; } = string.Empty;
        public string TotalWFHHours { get; set; } = string.Empty;
        public string ActualWFHHours { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string FirstHalfStatus { get; set; } = string.Empty;
        public string SecondHalfStatus { get; set; } = string.Empty;
    }
}
