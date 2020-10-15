using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace JdeScanExcelAddIn.Static
{
    public static class Functions
    {
        public static DateTime FirstDateOfWeek(int year, int week)
        {
                DateTime jan1 = new DateTime(year, 1, 1);
                int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

                // Use first Thursday in January to get first week of the year as
                // it will never be in Week 52/53
                DateTime firstThursday = jan1.AddDays(daysOffset);
                var cal = CultureInfo.CurrentCulture.Calendar;
                int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

                var weekNum = week;
                // As we're adding days to a date in Week 1,
                // we need to subtract 1 in order to get the right date for week #1
                if (firstWeek == 1)
                {
                    weekNum -= 1;
                }

                // Using the first Thursday as starting week ensures that we are starting in the right year
                // then we add number of weeks multiplied with days
                var result = firstThursday.AddDays(weekNum * 7);

                // Subtract 3 days from Thursday to get Monday, which is the first weekday in ISO8601
                return result.AddDays(-3);
        }

        public static void ConfigNlog()
        {
            string xml = @"
            <nlog xmlns=""http://www.nlog-project.org/schemas/NLog.xsd""
                xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
                <variable name=""logDirectory"" value=""C:/Inne/logs/ "" />
                < targets>
                    <target name=""file"" xsi:type=""File""
                            layout=""${longdate} ${logger} ${message}"" 
                            fileName=""${specialfolder:ApplicationData}\JdeScanExcelAddIn\log.txt""
                            keepFileOpen=""false""
                            encoding=""iso-8859-2"" />
                    <target xsi:type=""File"" name=""fileLog"" fileName=""${logDirectory}App.log""
                        layout=""${longdate} ${uppercase:${level}} ${message}"" />
                </targets>
                <rules>
                    <logger name=""*""  writeTo=""fileLog"" />
                </rules>
            </nlog>";

            StringReader sr = new StringReader(xml);
            XmlReader xr = XmlReader.Create(sr);
            XmlLoggingConfiguration config = new XmlLoggingConfiguration(xr, null);
            NLog.LogManager.Configuration = config;
        }

        //public static void ConfigNlog2()
        //{
        //    var target = new FileTarget
        //    {
        //        FileName = logfile,
        //        ReplaceFileContentsOnEachWrite = true,
        //        CreateDirs = createDirs
        //    };
        //    var config = new LoggingConfiguration();

        //    config.AddTarget("logfile", target);

        //    config.AddRuleForAllLevels(target);

        //    LogManager.Configuration = config;
        //}

        public static void ConfigNLog3()
        {
            var pathToNlogConfig = "c:\\Inne\\NLog.config";
            var config = new XmlLoggingConfiguration(pathToNlogConfig);
            LogManager.Configuration = config;
        }
    }
}
