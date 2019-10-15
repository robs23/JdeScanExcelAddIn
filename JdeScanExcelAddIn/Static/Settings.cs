using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JdeScanExcelAddIn.Models
{
    public static class Settings
    {
        public static int? Week { get; set; }
        public static int? Year { get; set; }
        public static DateTime? StartDate { get; set; }
        public static DateTime? EndDate { get; set; }

        private static SqlConnection _conn { get; set; }
        public static SqlConnection conn
        {
            get
            {
                if (_conn == null)
                {
                    _conn = new SqlConnection(Static.Secrets.ConnectionString);
                }
                if (_conn.State == System.Data.ConnectionState.Closed || _conn.State == System.Data.ConnectionState.Closed)
                {
                    try
                    {
                        _conn.Open();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Nie udało się nawiązać połączenia z bazą danych.. " + ex.Message);
                    }

                }
                return _conn;
            }
        }
    }
}
