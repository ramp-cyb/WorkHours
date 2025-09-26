namespace CybageMISAutomation.Constants
{
    /// <summary>
    /// Constants for web element IDs and selectors used in web automation.
    /// Centralizing these makes maintenance easier and reduces magic strings throughout the code.
    /// </summary>
    public static class WebElementConstants
    {
        // Element IDs
        public static class ElementIds
        {
            public const string SwipeLogLink = "TempleteTreeViewt32";
            public const string EmployeeDropdown = "EmployeeIDDropDownList7322";
            public const string DayDropdown = "DayDropDownList8665";
            public const string GenerateButton = "ViewReportImageButton";
            public const string ReportViewer = "ReportViewer1";
            public const string WizardPanel = "WizardPanel";
        }

        // CSS Selectors
        public static class Selectors
        {
            public const string GenerateButton = "input[name*=\"ViewReport\"], input[title*=\"Generate\"], input[value*=\"Generate\"]";
            public const string AlternativeGenerateButton = "input[type=\"submit\"], input[type=\"button\"][onclick*=\"Report\"]";
            public const string FallbackGenerateButton = "button[onclick*=\"Report\"], a[onclick*=\"Report\"]";
            public const string ViewShowButtons = "input[value*=\"View\"], input[value*=\"Show\"]";
            
            public const string ReportTable = "#ReportViewer1 table, table[id*=\"report\"], .report table, table.ReportTable";
            public const string ReportContent = "#ReportViewer1, div[id*=\"Report\"], iframe[src*=\"Report\"]";
            public const string ErrorMessages = ".error, .Error, div:contains(\"No Data\"), div:contains(\"Error\")";
            
            public const string EmployeeTable = "table:has(td:contains(\"Employee\")), table:has(th:contains(\"Employee\"))";
            public const string ReportDiv = "div[id*=\"Report\"], div[class*=\"report\"]";
            public const string ReportClass = ".report, .Report";
        }

        // Common text patterns for element identification
        public static class TextPatterns
        {
            public const string EmployeeKeyword = "Employee";
            public const string DateKeyword = "Date";
            public const string TimeKeyword = "Time";
            public const string DirectionKeyword = "Direction";
            public const string MachineKeyword = "Machine";
            public const string GenerateKeyword = "Generate";
            public const string ViewKeyword = "View";
            public const string ReportKeyword = "Report";
        }

        // Timeouts and delays (in milliseconds)
        public static class Timeouts
        {
            public const int DefaultWait = 2000;
            public const int NavigationWait = 3000;
            public const int ReportGenerationWait = 5000;
            public const int ElementSearchWait = 1000;
            public const int ScriptExecutionTimeout = 10000;
        }
    }
}