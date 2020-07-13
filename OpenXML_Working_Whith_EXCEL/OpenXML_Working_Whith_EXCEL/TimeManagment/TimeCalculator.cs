using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Working_Whith_EXCEL
{
    class TimeCalculator
    {/// <summary>
     /// Высчитывает параметр  AppealFinishingTime для элементов Appeal с учетом заданных стоимостей по времени
     /// </summary>
     /// <param name="appeal"></param>
     /// <param name="timeToFinish"></param>
     /// <returns></returns>
       static public Appeal FinishTimeCalculator(Appeal appeal ,SpendedTimeConfig timeToFinish)
        {
            switch(appeal.AppealCathegory)
            {
                case Appeal.AppealCathegoryEnum.Incident4LowPriority:
                    appeal.AppealFinishingTime = appeal.TakenInWork.AddMinutes(timeToFinish.Incident4LowPriorityTimeCosts);
                    break;

                case Appeal.AppealCathegoryEnum.Incident3MediumPriority:
                    appeal.AppealFinishingTime = appeal.TakenInWork.AddMinutes(timeToFinish.Incident3MediumPriorityTimeCosts);
                    break;

                case Appeal.AppealCathegoryEnum.Incident1ExtremeHighPriority:
                    appeal.AppealFinishingTime = appeal.TakenInWork.AddMinutes(timeToFinish.Incident1ExtremeHighPriorityTimeCosts);
                    break;
                case Appeal.AppealCathegoryEnum.RevisionUnknownPriority:
                    appeal.AppealFinishingTime = appeal.TakenInWork.AddMinutes(timeToFinish.RevisionUnknownPriorityTimeCosts);
                    break;
                case Appeal.AppealCathegoryEnum.ConsultationLowPriority:
                    appeal.AppealFinishingTime = appeal.TakenInWork.AddMinutes(timeToFinish.ConsultationLowPriorityTimeCosts);
                    break;
                case Appeal.AppealCathegoryEnum.StandartServiceRequest:
                    appeal.AppealFinishingTime = appeal.TakenInWork.AddMinutes(timeToFinish.StandartServiceRequestTimeCosts);
                    break;
                case Appeal.AppealCathegoryEnum.ConsultationMediumPriority:
                    appeal.AppealFinishingTime = appeal.TakenInWork.AddMinutes(timeToFinish.ConsultationMediumPriorityTimeCosts);
                    break;
                case Appeal.AppealCathegoryEnum.TuningServiceRequest:
                    appeal.AppealFinishingTime = appeal.TakenInWork.AddMinutes(timeToFinish.TuningServiceRequestTimeCosts);
                    break;
            }
            return appeal;
        }
    }
}
