using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Collections.Generic;
using System.Linq;
using System;
using CybageMISAutomation.Models;
// Converter is in the same namespace

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
            
            // Set window size to 60% of desktop
            var workArea = SystemParameters.WorkArea;
            Width = workArea.Width * 0.6;
            Height = workArea.Height * 0.6;
            
            // Center the window
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            
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
                        Text = day.Date.ToString("dd-MMM"),
                        FontWeight = FontWeights.Bold,
                        FontSize = 22,
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
                        FontSize = 18,
                        FontWeight = FontWeights.Medium,
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
            
            // Add weekly totals in column 7
            for (int row = 1; row <= 6; row++)
            {
                var weekStart = (row - 1) * 7;
                var weekEnd = Math.Min(weekStart + 7, allDays.Count);
                
                var weeklyTotal = 0.0;
                var weeklyDays = 0;
                
                for (int i = weekStart; i < weekEnd; i++)
                {
                    if (i < allDays.Count && !allDays[i].IsPlaceholder && allDays[i].HoursDecimal > 0)
                    {
                        weeklyTotal += allDays[i].HoursDecimal;
                        weeklyDays++;
                    }
                }
                
                var weeklyBorder = new Border();
                weeklyBorder.SetValue(Grid.RowProperty, row);
                weeklyBorder.SetValue(Grid.ColumnProperty, 7);
                weeklyBorder.Style = (Style)FindResource("CalendarCellStyle");
                
                // Apply color coding based on weekly hours: 40-45 green, 45+ blue, <40 red
                SolidColorBrush weeklyBackground;
                if (weeklyTotal >= 45)
                {
                    weeklyBackground = new SolidColorBrush(Color.FromRgb(0x87, 0xCE, 0xFA)); // Light blue
                }
                else if (weeklyTotal >= 40)
                {
                    weeklyBackground = new SolidColorBrush(Color.FromRgb(0x90, 0xEE, 0x90)); // Light green
                }
                else if (weeklyTotal > 0)
                {
                    weeklyBackground = new SolidColorBrush(Color.FromRgb(0xFF, 0xB6, 0xC1)); // Light red
                }
                else
                {
                    weeklyBackground = new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8)); // Light gray for no data
                }
                weeklyBorder.Background = weeklyBackground;
                
                var weeklyGrid = new Grid();
                weeklyGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                weeklyGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                
                var weekLabel = new TextBlock
                {
                    Text = $"Week {row}",
                    FontWeight = FontWeights.Bold,
                    FontSize = 14,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66)),
                    HorizontalAlignment = HorizontalAlignment.Center
                };
                Grid.SetRow(weekLabel, 0);
                weeklyGrid.Children.Add(weekLabel);
                
                var weeklyHours = new TextBlock
                {
                    Text = weeklyTotal > 0 ? $"{weeklyTotal:F1}h" : "-",
                    FontSize = 16,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x11, 0x11, 0x11)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetRow(weeklyHours, 1);
                weeklyGrid.Children.Add(weeklyHours);
                
                weeklyBorder.Child = weeklyGrid;
                CalendarGrid.Children.Add(weeklyBorder);
            }
        }
    }
}