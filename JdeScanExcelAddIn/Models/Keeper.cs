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

        public Keeper()
        {
            Items = new List<T>();
        }

        

    }
}
