using JdeScanExcelAddIn.Static;
using NLog;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JdeScanExcelAddIn.Models
{
    public class PlaceKeeper : Keeper<Place>
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public void Reload()
        {
            string sql = "SELECT PlaceId, Name, IsArchived, Priority FROM JDE_Places";

            SqlCommand sqlComand;
            sqlComand = new SqlCommand(sql, Settings.conn);
            using (SqlDataReader reader = sqlComand.ExecuteReader())
            {
                while (reader.Read())
                {
                    Place p = new Place { PlaceId = reader.GetInt32(reader.GetOrdinal("PlaceId")), Name = reader["Name"].ToString().Trim()};
                    p.IsArchived = reader.GetValueOrDefault<bool>("IsArchived");
                    p.Priority = reader.GetValueOrDefault<string>("Priority");
                    Items.Add(p);
                }
            }
        }

        public async Task<string> UpdatePriority()
        {
            Logger.Info("UpdatePriority has started.");
            List<Task<string>> UpdateTasks = new List<Task<string>>();

            try
            {
                foreach (Place p in Items.Where(i => i.IsUpdated==true))
                {
                    UpdateTasks.Add(Task.Run(()=> p.Edit()));
                }

                string response = "";

                IEnumerable<string> res = await Task.WhenAll<string>(UpdateTasks);
                if (res.Any())
                {
                    foreach (string r in res)
                    {
                        if (!string.IsNullOrEmpty(r))
                        {
                            if (r != "OK")
                            {
                                response += r;
                            }
                        }
                    }
                    if (string.IsNullOrWhiteSpace(response))
                        response = "OK";
                    Logger.Info("UpdatePriority has finished successfully.");
                    return response;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error in UpdatePriority. Error={ex}", ex);
                throw;
            }
            Logger.Info("UpdatePriority has finished unsuccessfully (no place updated)");
            return "Nie udało się zaktualizować danych żadnego zasobu..";
        }

    }
}
