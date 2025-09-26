namespace CybageMISAutomationLinux.Models
{
    public class MonthlyAttendanceEntry
    {
        public string EmployeeId { get; set; } = "";
        public string EmployeeName { get; set; } = "";
        public string Date { get; set; } = "";
        public int SwipeCount { get; set; }
        public string InTime { get; set; } = "";
        public string OutTime { get; set; } = "";
        public string TotalHours { get; set; } = "";
        public string ActualWorkHours { get; set; } = ""; // This is the "Actual Working Hours Swipe (A) + WFH (B) (HH:MM)" column
        public string TotalWFHHours { get; set; } = "";
        public string ActualWFHHours { get; set; } = "";
        public string Status { get; set; } = "";
        public string FirstHalfStatus { get; set; } = "";
        public string SecondHalfStatus { get; set; } = "";
    }
}