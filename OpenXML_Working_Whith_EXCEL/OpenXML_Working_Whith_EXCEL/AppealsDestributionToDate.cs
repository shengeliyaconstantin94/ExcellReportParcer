using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Working_Whith_EXCEL
{
    public class AppealsDestributionToDate
    {
        /// <summary>
        /// Метод высчитывает среднее количество жалоб в день
        /// </summary>
        /// <param name="appeals"></param>
        /// <param name="workDays"></param>
        /// <returns></returns>
        public static float AverageAppealsCount(List<Appeal> appeals, DateTime[] workDays)
        {
            float appealsCount = appeals.Count;
            float workDaysCount = workDays.Length;
            return appealsCount / workDaysCount;

        }

        public static List<Appeal> AppealsDistributionOnDate (List<Appeal> appeals , DateTime[] workDays)
        {

            
                for (int j = 0; j < appeals.Count; j++)
                    appeals[j].CreationDate = workDays[j % workDays.Length];

                return appeals;
            




        }
           
        }
            
        }
        

        
    

