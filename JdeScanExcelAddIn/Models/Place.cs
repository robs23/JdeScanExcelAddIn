using Microsoft.Office.Core;
using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JdeScanExcelAddIn.Models
{
    public class Place : Entity<Place>
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        [DisplayName("ID")]
        public int PlaceId { get; set; }
        public override int Id
        {
            set => value = PlaceId;
            get => PlaceId;
        }
        public string Name { get; set; }

        public bool? IsArchived { get; set; }
        public string Priority  { get; set; }
        public bool? IsUpdated { get; set; } = null;

        public async Task<string> Edit()
        {
            DateTime start = DateTime.Now;
            DateTime connDel = DateTime.Now;
            string iSql = @"UPDATE JDE_Places
                            SET Priority=@Priority
                            WHERE PlaceId=@PlaceId";
            string msg = "OK";

            DateTime connReq = DateTime.Now;
            using (SqlCommand command = new SqlCommand(iSql, Settings.conn))
            {
                connDel = DateTime.Now;
                command.Parameters.AddWithValue("@PlaceId", PlaceId);
                command.Parameters.AddWithValue("@Priority", Priority);

                int result = -1;

                try
                {
                    result = command.ExecuteNonQuery();
                    IsUpdated = false;
                }
                catch (Exception ex)
                {
                    Logger.Error("Error in Place of id={PlaceId}. Error={ex}.", PlaceId,ex);
                    msg = $"Wystąpił błąd przy edycji zasobu {Name}. Opis błędu: {ex.Message}";
                }

            }
            DateTime end = DateTime.Now;
            Logger.Info("Place of id={PlaceId} has finished in {TotalTime} sec. It waited {ConnTime} sec for connection.", PlaceId,(end-start).TotalSeconds, (connDel-connReq).TotalSeconds);
            return msg;
        }
    }
}
