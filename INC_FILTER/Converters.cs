using Microsoft.Office.Core;
using Newtonsoft.Json;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using static INC_FILTER.Settings;

namespace INC_FILTER
{
    public static class StaticData
    {
        static StaticData()
        {
            var settings = JsonConvert.DeserializeObject<Settings>(File.ReadAllText("settings.json"));
            HighlightSenders = settings.HighlightSenders;
            HighlightTrelloStatus = settings.HighlightTrelloStatus;

        }
        public static HighlightValue[] HighlightSenders;
        public static HighlightValue[] HighlightTrelloStatus;
    }


    public class IsSendToMeColorConverter : IValueConverter
    {
        

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null || !(value is IncEmailItem)) return new SolidColorBrush(Colors.Transparent);
            var senderName = ((IncEmailItem)value).SenderName;

            foreach (var highlightSender in StaticData.HighlightSenders)
            {

                if (highlightSender.Values.Any(x => senderName.ToLower().Contains(x.ToLower())))
                {
                    return new SolidColorBrush(Color.FromRgb(highlightSender.Color[0], highlightSender.Color[1], highlightSender.Color[2]));
                }
            }
            return new SolidColorBrush(Colors.Transparent);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class IsActiveListColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null) return new SolidColorBrush(Colors.Transparent);
            if (value is string)
            {
                var t = (string)value;

                foreach (var highlightStatus in StaticData.HighlightTrelloStatus)
                {
                    if (highlightStatus.Values.Any(x => t.ToLower().Contains(x.ToLower())))
                        return new SolidColorBrush(Color.FromRgb(highlightStatus.Color[0], highlightStatus.Color[1], highlightStatus.Color[2]));
                }


            }
            return new SolidColorBrush(Colors.Transparent);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class CollapsedWhenNullConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null) return Visibility.Collapsed;
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class TrelloNameTrimmerConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null) return string.Empty;
            var x = (string)value;
            x = x.Trim();
            if (x.StartsWith("ICT_"))
            {
                return x.Substring(15, x.Length-15);
            }

            return x;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class DisplayDateTimeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is null) return string.Empty;
            var x = (DateTime)value;
           

            return x.ToString("HH:mm dd MMM") + " | " + GetDiff(x);
        }

        private string GetDiff(DateTime x)
        {
            var totalHours = Math.Floor((DateTime.Now - x).TotalHours);

            var totalDay = Math.Floor(totalHours / 24);
            var totalHour = totalHours % 24;
            if ((DateTime.Now - x).TotalMinutes < 59)
            {
                return Math.Floor((DateTime.Now - x).TotalMinutes) + "m ago";
            }
            if (totalDay <= 0.001) return totalHour + "h ago"; else return x.ToString("ddd") + ", " + totalDay + "d ago";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

}
