using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace OpenXML_Working_Whith_EXCEL
{
   public class WorkBookConfig
    {
        public SpreadsheetDocument SpredSheet { get; set; }
        public string WorkSheetName { get; set; }
        public string PathToBook { get; set; }
    }
}
