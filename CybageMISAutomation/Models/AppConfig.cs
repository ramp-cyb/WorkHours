using System;

namespace CybageMISAutomation.Models
{
    public class AppConfig
    {
        public string EmployeeId { get; set; } = "1476";
        public string MisUrl { get; set; } = "https://cybagemis.cybage.com/Report%20Builder/RPTN/ReportPage.aspx";
        public bool ShowLogWindow { get; set; } = false;
        public bool ShowMonthly { get; set; } = false;
        public int AutomationDelayMs { get; set; } = 2000;
        public string WindowTitle { get; set; } = "Cybage MIS Report Automation";
    }
}