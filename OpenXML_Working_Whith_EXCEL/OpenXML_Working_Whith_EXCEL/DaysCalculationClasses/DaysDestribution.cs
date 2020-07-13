using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using NLog.Common;
using System.Xml;
using System.Globalization;
using System.Linq;

namespace OpenXML_Working_Whith_EXCEL
{
    class DaysDestribution
    {
        
      
       

        /// <summary>
        /// Метод для отладки приложения
        /// </summary>
        /// <param name="config">файл,содержащий в свойствах необходимые параметры(дата начала квартала,дата конца квартала, коллекцию выходных дней)</param>
        internal static void Writer(Config config)
        {
            DateTime[] workDaysOnFirstQuater = WorkDayCalculator(config.StartDate, config.EndDate, config.HoliDays);
            foreach (DateTime s in workDaysOnFirstQuater)
            {
                Console.WriteLine(s.ToString());
            }
        }
        /// <summary>
        ///  метод для конвертирования массива строк в массив дат
        /// </summary>
        /// <param name="holiday"></param>
        /// <returns></returns>
        static internal DateTime[] HolidaysConverter(string[] holiday)  

        {
            DateTime[] holidaysInDataTime = new DateTime[holiday.Length];
            CultureInfo cult = CultureInfo.CurrentCulture;
            for (int i = 0; i <= holiday.Length-1; i++)
            {
                holidaysInDataTime[i] = DateTime.Parse(holiday[i]); 
            }

            return holidaysInDataTime;
        }
        /// <summary>
        /// Расчет рабочих дней путем проверки вхождения множества
        /// </summary>
        /// <param name="quaterStartDate"></param>
        /// <param name="quaterEndDate"></param>
        /// <param name="holidaysInDateTime">Принимает любую "коллекцию" типа DateTime и реализующую IEnumerable</param>
        /// <returns></returns>
        static internal DateTime[] WorkDayCalculator (DateTime quaterStartDate, DateTime quaterEndDate, IEnumerable<DateTime> holidaysInDateTime)
        {
            TimeSpan time = quaterEndDate - quaterStartDate;
            DateTime[] rawDayCount = new DateTime[time.Days];
            for (int i=0;i<=(time.Days -1); i++)
            {
                rawDayCount[i] = quaterStartDate.AddDays(i);

            }
            var WorkDays = rawDayCount.Except(holidaysInDateTime);
            return WorkDays.ToArray<DateTime>();

                }
    }
}
