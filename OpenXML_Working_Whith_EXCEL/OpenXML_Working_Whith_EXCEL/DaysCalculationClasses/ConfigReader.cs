using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Working_Whith_EXCEL
{
    public class ConfigReader
    {
    /// <summary>
    /// Метод для чтения из файла конфигурации
    /// </summary>
    /// <returns></returns>
        static public Config ReadConfig ()
        {
             char spliter = Convert.ToChar(ConfigurationManager.AppSettings["spliter"]);


             DateTime firstQuaterStartDate = Convert.ToDateTime(ConfigurationManager.AppSettings["secondQuarterStartDate"]);
             DateTime firstQuaterEndDate = Convert.ToDateTime(ConfigurationManager.AppSettings["secondQuarterEndDate"]);
            

             string[] holiday = ConfigurationManager.AppSettings["holidays"].Split(spliter);
             DateTime[] holidaysInDateTime = DaysDestribution.HolidaysConverter(holiday);

            return new Config() { EndDate = firstQuaterEndDate, StartDate = firstQuaterStartDate, HoliDays = holidaysInDateTime };
        }
    }
}
