using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Working_Whith_EXCEL
{
    class TimeCostsGenerator
    {
        /// <summary>
        /// Метод генерирует экземпляр SpendedTimeConfig, содержащий в свойствах временные затраты на закрытие обращений 
        /// </summary>
        /// <returns></returns>
       static public SpendedTimeConfig TimeCostsGeneration()
        {
            Random rnd = new Random();
            int consLowPrior = rnd.Next(30, 180);
            int consMedPrior = rnd.Next(30, 120);
            int incExHiPrior = rnd.Next(30, 90);
            int incMedPrior = rnd.Next(30, 120);
            int incLowPrior = rnd.Next(60, 120);
            int revUnkPrior = rnd.Next(60, 180);
            int standServReq = rnd.Next(60, 480);
            int tuneServReq = rnd.Next(30, 120);
            return new SpendedTimeConfig
            {
                ConsultationLowPriorityTimeCosts = consLowPrior,
                ConsultationMediumPriorityTimeCosts = consMedPrior,
                Incident1ExtremeHighPriorityTimeCosts = incExHiPrior,
                Incident3MediumPriorityTimeCosts = incMedPrior,
                Incident4LowPriorityTimeCosts = incLowPrior,
                RevisionUnknownPriorityTimeCosts = revUnkPrior,
                StandartServiceRequestTimeCosts = standServReq,
                TuningServiceRequestTimeCosts = tuneServReq
            };
        }
    }
}
