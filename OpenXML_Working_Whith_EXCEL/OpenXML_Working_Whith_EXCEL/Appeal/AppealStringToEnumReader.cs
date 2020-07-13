using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Working_Whith_EXCEL
{
    class AppealStringToEnumReader
    {
        /// <summary>
        /// Конвертирует строковые значение в энум
        /// </summary>
        /// <param name="cathegoryString"></param>
        /// <returns></returns>
       static public Appeal.AppealCathegoryEnum ConvertToAppealEnum(string cathegoryString )
        {
            Appeal.AppealCathegoryEnum cathegoryEnum = Appeal.AppealCathegoryEnum.StandartServiceRequest;
            switch (cathegoryString)
            {
                case "Инцидент 4 приоритета (низкий)":
                    cathegoryEnum = Appeal.AppealCathegoryEnum.Incident4LowPriority;
                    break;
                case "Инцидент 3 приоритета (средний)":
                     cathegoryEnum = Appeal.AppealCathegoryEnum.Incident3MediumPriority;
                    break;
                case "Доработка":
                     cathegoryEnum = Appeal.AppealCathegoryEnum.RevisionUnknownPriority;
                    break;
                case "Консультация низкого приоритета":
                     cathegoryEnum = Appeal.AppealCathegoryEnum.ConsultationLowPriority;
                    break;
                case "ЗНО типовой":
                     cathegoryEnum = Appeal.AppealCathegoryEnum.StandartServiceRequest;
                    break;
                case "Консультация - работа в СЭДО":
                     cathegoryEnum = Appeal.AppealCathegoryEnum.ConsultationMediumPriority;
                    break;
                case "Инцидент 1 приоритета (очень высокий)":
                     cathegoryEnum = Appeal.AppealCathegoryEnum.Incident1ExtremeHighPriority;
                    break;
                case "ЗНО - настройка":
                     cathegoryEnum = Appeal.AppealCathegoryEnum.TuningServiceRequest;
                    break;


            }

            return cathegoryEnum;
        }
    }
}
