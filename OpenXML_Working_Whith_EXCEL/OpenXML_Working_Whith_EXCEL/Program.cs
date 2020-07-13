using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Configuration;
using System.Data;
using System.IO;
using OpenXML_Working_Whith_EXCEL.WorkWithExcelClasses;

namespace OpenXML_Working_Whith_EXCEL
{
    class Program
    {
        static void Main(string[] args)
        {
            
            WorkBookConfig s = WorkBookConfigReader.ReadConfig();
            List<object> sameList = WorkBookReader.ToList(s);
            List<Appeal> appeals = AppealConverter.ObjectListToAppeal(sameList);
            Config config = ConfigReader.ReadConfig();
           DateTime[] workDays = DaysDestribution.WorkDayCalculator(config.StartDate, config.EndDate, config.HoliDays);
            AppealsDestributionToDate.AppealsDistributionOnDate(appeals, workDays);
            ExportToWorkbook.ExportToExcell(appeals);
                Console.WriteLine("Рекомбинация выполнена успешно");
                
            
          
            Console.ReadLine();
            
            
        }
    }
}
