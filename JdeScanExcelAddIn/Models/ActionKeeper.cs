using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JdeScanExcelAddIn.Models
{
    public class ActionKeeper : Keeper<Action>
    {

        public void Reload()
        {
            string sql = "SELECT ActionId, Name FROM JDE_Actions";

            SqlCommand sqlComand;
            sqlComand = new SqlCommand(sql, Settings.conn);
            using (SqlDataReader reader = sqlComand.ExecuteReader())
            {
                while (reader.Read())
                {
                    Action a = new Action { ActionId = reader.GetInt32(reader.GetOrdinal("ActionId")), Name = reader["Name"].ToString().Trim()};
                    Items.Add(a);
                }
            }
        }
    }
}
