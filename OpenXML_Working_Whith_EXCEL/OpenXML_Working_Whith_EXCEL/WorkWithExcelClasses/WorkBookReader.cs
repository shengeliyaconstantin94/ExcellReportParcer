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
using ClosedXML.Excel;

namespace OpenXML_Working_Whith_EXCEL
{

    public class WorkBookReader
    {

        public static void LoadAndSaveFile(String input, String output)
        {
            var wb = new XLWorkbook(input);
            wb.SaveAs(output);
            wb.SaveAs(output);
        }


        static public List<object> ToList(WorkBookConfig config)
        {
            List<object> dataFromTable = new List<object>();

            using (var excelWorkbook = new XLWorkbook(config.PathToBook))
            {
                var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();

                foreach (var dataRow in nonEmptyDataRows)
                {

                    for (int i = 1; i <= 7; i++)
                    {
                        if ((i == 4 || i == 5) && (dataRow.RangeAddress.FirstAddress.RowNumber !=1))//Проверяем, не пренадлежат ли данные к первой строк(так как в ней заложены названия столбцов)
                        {
                            if (dataRow.Cell(i).DataType == XLDataType.Number) //Проверяем, является ли содержимое ячейки цифрой (так как даты интерпретируются как значени double)
                            {
                                object test = dataRow.Cell(i).Value;


                                double dobTest = (double)test;
                                DateTime dtTest = DateTime.FromOADate(dobTest);
                                dataFromTable.Add(dtTest);
                            }
                            
                        }
                        else
                        {
                            dataFromTable.Add(dataRow.Cell(i).Value);
                        }
                    }
                }
            }
            return dataFromTable;

        }
    }
}
