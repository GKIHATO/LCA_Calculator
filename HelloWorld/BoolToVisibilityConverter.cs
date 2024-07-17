using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows;

namespace HelloWorld
{
    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue && targetType == typeof(Visibility))
            {
                return boolValue ? Visibility.Visible : Visibility.Collapsed;
                // Alternatively, you can use Visibility.Hidden instead of Visibility.Collapsed if you want to hide without collapsing the space.
            }

            return DependencyProperty.UnsetValue; // Return DependencyProperty.UnsetValue for unsupported conversions.
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException(); // ConvertBack is not needed for this one-way conversion.
        }
    }
}
