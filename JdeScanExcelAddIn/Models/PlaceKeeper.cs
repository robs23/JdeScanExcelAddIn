using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JdeScanExcelAddIn.Models
{
    public class PlaceKeeper : Keeper<Place>
    {
        public void Reload()
        {
            string sql = "SELECT PlaceId, Name FROM JDE_Places";

            SqlCommand sqlComand;
            sqlComand = new SqlCommand(sql, Settings.conn);
            using (SqlDataReader reader = sqlComand.ExecuteReader())
            {
                while (reader.Read())
                {
                    Place p = new Place { PlaceId = reader.GetInt32(reader.GetOrdinal("PlaceId")), Name = reader["Name"].ToString().Trim()};
                    Items.Add(p);
                }
            }
        }
    }
}
