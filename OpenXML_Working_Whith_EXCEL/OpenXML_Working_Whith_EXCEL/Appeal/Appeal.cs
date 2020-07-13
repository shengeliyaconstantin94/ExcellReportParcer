using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Working_Whith_EXCEL
{
   public class Appeal
    {
       public enum AppealCathegoryEnum
        {
            Incident1ExtremeHighPriority,
            Incident2HighPriority,
            Incident3MediumPriority,
            Incident4LowPriority,
            RevisionHighPriority,
            RevisionMediumPriority,
            RevisionLowPiority,
            RevisionUnknownPriority,
            RevisionRateHighPriority,
            RevisionRateMediumPriority,
            RevisionRateLowPriority,
            SolutionErrorHighPriority,
            SolutionErrorMediumPriority,
            SolutionErrorLowPriority,
            UserErrorHighPriority,
            UserErrorMediumPriority,
            UserErrorLowPiority,
            PlatformErrorHighPriority,
            PlatformErrorMediumPriority,
            PlatformErrorLowPriority,
            MalfunctionHighPriority,
            MalfunctionMediumPriority,
            MalfunctionLowPriority,
            ConsultationHighPriority,
            ConsultationMediumPriority,
            ConsultationLowPriority,
            ConsultationAndTuning,
            StandartServiceRequest,
            TuningServiceRequest,
        }
        public string AppealNumber { get; set; }
        public string AppealTheme { get; set; }
        public string AppealInitiator { get; set; }
        public DateTime CreationDate { get; set; }
        public DateTime TakenInWork { get; set; }
        public string AppealPerformer { get; set; }
        public string AppealsCathegoryString { get; set; }
        public AppealCathegoryEnum AppealCathegory { get; set; }
        public DateTime AppealFinishingTime { get; set; }



    }
}
