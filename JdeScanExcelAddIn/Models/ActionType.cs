using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JdeScanExcelAddIn.Models
{
    public class ActionType: Entity<ActionType>
    {
        public int ActionTypeId { get; set; }
        public override int Id
        {
            set => value = ActionTypeId;
            get => ActionTypeId;
        }
        public string Name { get; set; }
        public string Description { get; set; }
        public bool? RequireInitialDiagnosis { get; set; }
        public bool? ShowInPlanning { get; set; }
        public bool? MesSync { get; set; }
        public bool? AllowDuplicates { get; set; }
        public bool? RequireQrToStart { get; set; }
        public bool? RequireQrToFinish { get; set; }
        public bool? ClosePreviousInSamePlace { get; set; }
        public bool? PartsApplicable { get; set; }
        public bool? ActionsApplicable { get; set; }
    }
}
