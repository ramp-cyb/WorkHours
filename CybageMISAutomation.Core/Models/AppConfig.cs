using System;

namespace CybageMISAutomation.Models
{
    public class AppConfig
    {
        public string EmployeeId { get; set; } = "";
        public string MisUrl { get; set; } = "https://cybagemis.cybage.com/Report%20Builder/RPTN/ReportPage.aspx";
        public bool ShowLogWindow { get; set; } = false;
        public bool ShowMonthly { get; set; } = false;
        public bool AutoStartFullReport { get; set; } = true;
        public int AutomationDelayMs { get; set; } = 2000;
        public string WindowTitle { get; set; } = "Cybage MIS Report Automation";
    }
}
