using JdeScanExcelAddIn.Static;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JdeScanExcelAddIn.Models
{
    public abstract class Entity<T>
    {
        [Browsable(false)]
        public abstract int Id { get; set; }
        [Browsable(false)]
        public int CreatedBy { get; set; }
        [DisplayName("Utworzył")]
        public string CreatedByName { get; set; }
        [DisplayName("Data utworzenia")]
        public DateTime CreatedOn { get; set; }
        [Browsable(false)]
        public int? LmBy { get; set; }
        [DisplayName("Zmodyfikował")]
        public string LmByName { get; set; }
        [DisplayName("Data modyfikacji")]
        public DateTime? LmOn { get; set; }
        [Browsable(false)]
        public int TenantId { get; set; }
        [Browsable(false)]
        public string TenantName { get; set; }
        [Browsable(false)]
        public int AddedId { get; set; }

        

    }
}
