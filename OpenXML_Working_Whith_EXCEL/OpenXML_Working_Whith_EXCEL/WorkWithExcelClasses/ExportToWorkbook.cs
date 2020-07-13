using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Working_Whith_EXCEL.WorkWithExcelClasses
{
    class ExportToWorkbook
    {
        public static void  ExportToExcell(List<Appeal> appealsToExcellData)
        {

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Лист 1");

            ws.Cell("A1").Value = "Номер обращения";
            ws.Cell("B1").Value = "Тема";
            ws.Cell("C1").Value = "Инициатор";
            ws.Cell("D1").Value = "Дата создания";
            ws.Cell("E1").Value = "Дата взятия в работу";
            ws.Cell("F1").Value = "Исполнитель";
            ws.Cell("G1").Value = "Категория";
            ws.Cell("H1").Value = "Норматив решения";

            for (int i = 0; i < appealsToExcellData.Count; i++)
            {
                string appealNumber = appealsToExcellData[i].AppealNumber;
                string appealTheme = appealsToExcellData[i].AppealTheme;
                string appealInitiator = appealsToExcellData[i].AppealInitiator;
                DateTime appealCreationDate = appealsToExcellData[i].CreationDate;
                DateTime appealsTakenInWork = appealsToExcellData[i].TakenInWork;
                string appealPerformer = appealsToExcellData[i].AppealPerformer;
                string appealCathegory = appealsToExcellData[i].AppealsCathegoryString;
                DateTime appealFinishingTime = appealsToExcellData[i].AppealFinishingTime;

                ws.Cell("A" + (i + 2)).Value = appealNumber;
                ws.Cell("B" + (i + 2)).Value = appealTheme;
                ws.Cell("C" + (i + 2)).Value = appealInitiator;
                ws.Cell("D" + (i + 2)).Value = appealCreationDate;
                ws.Cell("E" + (i + 2)).Value = appealsTakenInWork;
                ws.Cell("F" + (i + 2)).Value = appealPerformer;
                ws.Cell("G" + (i + 2)).Value = appealCathegory;
                ws.Cell("H" + (i + 2)).Value = appealFinishingTime;
                
            }

            // Beautify
            ws.Range("A1:H1").Style.Font.FontColor = XLColor.White ;
            ws.Range("A1:H1").Style.Fill.BackgroundColor = XLColor.Blue;
            ws.Columns().AdjustToContents();

            wb.SaveAs(ConfigurationManager.AppSettings["pathToSaveWorkbook"]);
        }
    }
}
