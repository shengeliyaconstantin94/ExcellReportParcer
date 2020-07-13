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

namespace OpenXML_Working_Whith_EXCEL.WorkWithExcelClasses
{
    public class WorkBookConfigReader
    {
        static public WorkBookConfig ReadConfig()
        {
            SpreadsheetDocument newSpredSheet = SpreadsheetDocument.Open(ConfigurationManager.AppSettings["pathToWorkbook"], false); // перенестив в WBC
             string workSheetName = ConfigurationManager.AppSettings["workSheetName"];
            string pathToWorkBook = ConfigurationManager.AppSettings["pathToWorkbook"];

            return new WorkBookConfig() { SpredSheet = newSpredSheet, WorkSheetName = workSheetName,PathToBook= pathToWorkBook };
        }
    }
}
