using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Working_Whith_EXCEL
{
    /// <summary>
    /// Класс-оболочка для  хранения и передачи параметров
    /// </summary>
   public class Config
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public DateTime [] HoliDays { get; set; }
    }
}
