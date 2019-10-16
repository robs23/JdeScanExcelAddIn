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
        //private SqlConnection _conn { get; set; }
        //public SqlConnection conn
        //{
        //    get
        //    {
        //        if(_conn == null)
        //        {
        //           _conn = new SqlConnection(Static.Secrets.ConnectionString);
        //        }
        //        if(_conn.State == System.Data.ConnectionState.Closed || _conn.State == System.Data.ConnectionState.Closed)
        //        {
        //            try
        //            {
        //                _conn.Open();
        //            }catch(Exception ex)
        //            {
        //                MessageBox.Show("Nie udało się nawiązać połączenia z bazą danych.. " + ex.Message);
        //            }
                    
        //        }
        //        return _conn;
        //    }
        //}

        public List<T> Items { get; set; } 

        public Keeper()
        {
            Items = new List<T>();
        }

        

    }
}
