using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Working_Whith_EXCEL
{
    public class AppealConverter
    {
        static public List<Appeal> ObjectListToAppeal(List<object> rawData)
        {
            List<Appeal> appealsList = new List<Appeal>();
            for (int i = 7; i <= rawData.Count-1; i += 7)
            {
                if (i + 3 < rawData.Count)
                {
                    if ((rawData[i + 3] is DateTime) && (rawData[i + 4] is DateTime))
                    {
                        appealsList.Add(new Appeal
                        {
                            AppealNumber = rawData[i].ToString(),
                            AppealTheme = rawData[i + 1].ToString(),
                            AppealInitiator = rawData[i + 2].ToString(),
                            CreationDate = Convert.ToDateTime(rawData[i + 3]),
                            TakenInWork = Convert.ToDateTime(rawData[i + 4]),
                            AppealPerformer = rawData[i + 5].ToString(),
                            AppealsCathegoryString = rawData[i + 6].ToString(),
                            AppealCathegory = AppealStringToEnumReader.ConvertToAppealEnum(rawData[i + 6].ToString()),
                            
                        });
                    }
                }
              
            }
            for (int i = 0; i <= appealsList.Count-1; i++)
            {
                
                appealsList[i] = TimeCalculator.FinishTimeCalculator(appealsList[i], TimeCostsGenerator.TimeCostsGeneration());
                WorkDayBorders dayBorders = WorkDayBorders.WorkDayBordersGenerator(appealsList[i].TakenInWork);
                if ((appealsList[i].AppealFinishingTime.Hour > dayBorders.EndOfWorkDay.Hour) ||
                    (appealsList[i].AppealFinishingTime.Hour == dayBorders.EndOfWorkDay.Hour && 
                    appealsList[i].AppealFinishingTime.Minute > dayBorders.EndOfWorkDay.Minute)  )
                {
                    
                    appealsList[i].AppealFinishingTime = new DateTime(appealsList[i].AppealFinishingTime.Year, appealsList[i].AppealFinishingTime.Month,
                        appealsList[i].AppealFinishingTime.AddDays(1).Day, dayBorders.StartOfWorkDay.Hour, 0, 0);
 
                }
                else if ((appealsList[i].AppealFinishingTime.Day > dayBorders.EndOfWorkDay.Day) && (appealsList[i].AppealFinishingTime.Hour < dayBorders.EndOfWorkDay.Hour))
                    {
                    appealsList[i].AppealFinishingTime = new DateTime(appealsList[i].AppealFinishingTime.Year, appealsList[i].AppealFinishingTime.Month,
                       appealsList[i].AppealFinishingTime.AddDays(1).Day, dayBorders.StartOfWorkDay.Hour, 0, 0);

                }
            }
            return appealsList;
                
            }
        }
    }
