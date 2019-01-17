using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace OutlookReportsAddIn
{
    public class BoolToImageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            bool val = (bool)value;
            if (val)
            {
                return new BitmapImage(new Uri(@"pack://application:,,,/OutlookReportsAddIn;component/Resources/check.png"));
                //return new BitmapImage(new Uri(@"\Resources\check.png", UriKind.Relative));
            }
            else
            {
                return new BitmapImage(new Uri(@"pack://application:,,,/OutlookReportsAddIn;component/Resources/error.png"));
                //return new BitmapImage(new Uri(@"Resources/error.png", UriKind.RelativeOrAbsolute));
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
