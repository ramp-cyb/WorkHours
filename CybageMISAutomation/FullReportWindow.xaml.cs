using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CybageMISAutomation.Models;

namespace CybageMISAutomation
{
    public partial class FullReportWindow : Window
    {
        private readonly FullReportViewModel _model;
        
        public FullReportWindow(FullReportViewModel model)
        {
            InitializeComponent();
            _model = model;
            DataContext = model;
            TxtSummary.Text = $"{model.Year}-{model.Month:00}  Worked: {model.WorkedDays}  Leave: {model.LeaveDays}  WO: {model.WeeklyOffDays}  Holidays: {model.HolidayDays}  Total Hrs: {model.TotalActualHours:F2}  Avg Hrs: {model.AverageActualHours:F2}";
            PopulateCalendarGrid();
        }
        
        private void PopulateCalendarGrid()
        {
            var converter = new HoursToBrushConverter();
            
            // Get all days from all weeks
            var allDays = new List<DayCalendarItem>();
            foreach (var week in _model.Weeks)
            {
                allDays.AddRange(week.Days);
            }
            
            // Populate grid cells (skip row 0 which is headers)
            for (int i = 0; i < allDays.Count && i < 42; i++) // Max 6 weeks * 7 days = 42
            {
                var day = allDays[i];
                int row = (i / 7) + 1; // +1 to skip header row
                int col = i % 7;
                
                if (row > 6) break; // Don't exceed our 6 data rows
                
                var border = new Border();
                border.SetValue(Grid.RowProperty, row);
                border.SetValue(Grid.ColumnProperty, col);
                border.Style = (Style)FindResource("CalendarCellStyle");
                border.ToolTip = day.Tooltip;
                
                var grid = new Grid();
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                
                if (!day.IsPlaceholder)
                {
                    // Day number + today indicator
                    var headerPanel = new StackPanel { Orientation = Orientation.Horizontal };
                    var dayText = new TextBlock
                    {
                        Text = day.Date.Day.ToString(),
                        FontWeight = FontWeights.Bold,
                        FontSize = 18,
                        Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
                    };
                    headerPanel.Children.Add(dayText);
                    
                    if (day.IsToday)
                    {
                        var todayIndicator = new System.Windows.Shapes.Ellipse
                        {
                            Width = 12,
                            Height = 12,
                            Margin = new Thickness(8, 0, 0, 0),
                            Fill = new SolidColorBrush(Color.FromRgb(0x3A, 0x7A, 0xFE))
                        };
                        headerPanel.Children.Add(todayIndicator);
                    }
                    Grid.SetRow(headerPanel, 0);
                    grid.Children.Add(headerPanel);
                    
                    // Hours display
                    var hoursText = new TextBlock
                    {
                        Text = day.Hours,
                        FontSize = 24,
                        FontWeight = FontWeights.SemiBold,
                        Foreground = new SolidColorBrush(Color.FromRgb(0x11, 0x11, 0x11)),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    Grid.SetRow(hoursText, 1);
                    grid.Children.Add(hoursText);
                    
                    // Apply hour-based background color to entire cell
                    var hoursBrush = (Brush?)converter.Convert(new object[] { day.HoursDecimal, day.Status }, typeof(Brush), null!, System.Globalization.CultureInfo.CurrentCulture);
                    if (hoursBrush != null && hoursBrush != Brushes.Transparent)
                    {
                        border.Background = hoursBrush;
                        border.Opacity = 0.7; // Make it subtle so text remains readable
                    }
                }
                
                border.Child = grid;
                CalendarGrid.Children.Add(border);
            }
        }
    }
}
