using JdeScanExcelAddIn.Static;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JdeScanExcelAddIn.Models
{
    public class ActionTypeKeeper: Keeper<ActionType>
    {
        public void Reload()
        {
            string sql = "SELECT ActionTypeId, Name FROM JDE_ActionTypes WHERE ShowInPlanning = 1";

            SqlCommand sqlComand;
            sqlComand = new SqlCommand(sql, Settings.conn);
            using (SqlDataReader reader = sqlComand.ExecuteReader())
            {
                while (reader.Read())
                {
                    ActionType at = new ActionType { ActionTypeId = reader.GetInt32(reader.GetOrdinal("ActionTypeId")), Name = reader["Name"].ToString().Trim()};
                    Items.Add(at);
                }
            }
        }
    }
}
