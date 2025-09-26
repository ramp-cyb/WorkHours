using System;

namespace CybageMISAutomationLinux.Models
{
	// Minimal representation for daily calculation aggregation.
	// Extend if needed to include raw swipe logs, etc.
	public class WorkHoursCalculation
	{
		public DateTime Date { get; set; }
		public string WorkingHoursDisplay { get; set; } = string.Empty; // e.g. calculated core working
		public string TotalWorkingHoursDisplay { get; set; } = string.Empty; // fallback
		public string Status { get; set; } = string.Empty; // optional daily derived status
	}
}
