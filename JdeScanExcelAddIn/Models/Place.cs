using Microsoft.Office.Core;
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
            string iSql = @"UPDATE JDE_Places
                            SET Priority=@Priority
                            WHERE PlaceId=@PlaceId";
            string msg = "OK";


            using (SqlCommand command = new SqlCommand(iSql, Settings.conn))
            {
                command.Parameters.AddWithValue("@PlaceId", PlaceId);
                command.Parameters.AddWithValue("@Priority", Priority);

                int result = -1;

                try
                {
                    result = await command.ExecuteNonQueryAsync();
                    IsUpdated = false;
                }
                catch (Exception ex)
                {
                    msg = $"Wystąpił błąd przy edycji zasobu {Name}. Opis błędu: {ex.Message}";
                }

            }
            return msg;
        }
    }
}
