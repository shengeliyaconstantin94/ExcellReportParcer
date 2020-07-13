using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Working_Whith_EXCEL
{
    class WorkDayBorders
    {
        public DateTime StartOfWorkDay { get; set; }
        public DateTime EndOfWorkDay { get; set; }
        
       /// <summary>
       /// Метод определяет рамки рабочего дня
       /// </summary>
       /// <param name="regularDay"></param>
       /// <returns></returns>
      static public WorkDayBorders WorkDayBordersGenerator (DateTime regularDay)
        {
            return new WorkDayBorders
            {
                StartOfWorkDay = new DateTime(regularDay.Year, regularDay.Month, regularDay.Day, 10, 0, 0),
                EndOfWorkDay = new DateTime(regularDay.Year, regularDay.Month, regularDay.Day, 20, 0, 0)
            };
        }
    }
}
