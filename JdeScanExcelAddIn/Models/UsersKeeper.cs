using JdeScanExcelAddIn.Static;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JdeScanExcelAddIn.Models
{
    public class UsersKeeper : Keeper<User>
    {

        public void Reload()
        {
            string sql = "SELECT UserId, Name, Surname, Password, IsArchived FROM JDE_Users";

            SqlCommand sqlComand;
            sqlComand = new SqlCommand(sql, Settings.conn);
            using (SqlDataReader reader = sqlComand.ExecuteReader())
            {
                while (reader.Read())
                {
                    User u = new User { UserId = reader.GetInt32(reader.GetOrdinal("UserId")), Name = reader["Name"].ToString().Trim(), Surname = reader["Surname"].ToString().Trim(), Password = reader["Password"].ToString()};
                    u.IsArchived = reader.GetValueOrDefault<bool>("IsArchived");
                    Items.Add(u);
                }
            }
        }
    }
}
