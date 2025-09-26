using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace CybageMISAutomation
{
    // Converts HoursDecimal (double) + Status to a background brush bar.
    // If status is not Present, we return Transparent so status coloring from cell style remains.
    public class HoursToBrushConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                if (values.Length >= 2 && values[0] is double hours && values[1] is string status)
                {
                    if (!string.Equals(status, "Present", StringComparison.OrdinalIgnoreCase))
                        return Brushes.Transparent;
                    if (hours <= 0) return Brushes.Transparent;
                    if (hours < 7) return new SolidColorBrush(Color.FromRgb(0xFF, 0xB8, 0xB8));
                    if (hours < 8) return new SolidColorBrush(Color.FromRgb(0xDF, 0xF5, 0xDD));
                    if (hours < 9) return new SolidColorBrush(Color.FromRgb(0xA6, 0xE3, 0xA1));
                    return new SolidColorBrush(Color.FromRgb(0x6E, 0xA8, 0xFF));
                }
            }
            catch (Exception ex)
            {
                // Log the error for debugging - silent failures make troubleshooting difficult
                System.Diagnostics.Debug.WriteLine($"HoursToBrushConverter error: {ex.Message}");
                return Brushes.Transparent;
            }
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }
}
