using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JdeScanExcelAddIn.Models
{

    public abstract class Keeper<T> 
    {
        public SqlConnection conn = new SqlConnection(Static.Secrets.ConnectionString);
        public List<T> Items { get; set; } 
        protected abstract string ObjectName { get;}
        protected abstract string PluralizedObjectName { get;}

        public Keeper()
        {
            Items = new List<T>();
        }

    }
}
