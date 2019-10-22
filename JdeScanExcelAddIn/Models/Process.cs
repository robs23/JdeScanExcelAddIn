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
    public class Process : Entity<Process>
    {
        [DisplayName("ID")]
        public int ProcessId { get; set; }
        public override int Id
        {
            set => value = ProcessId;
            get => ProcessId;
        }
        public Place Place { get; set; }
        public DateTime? PlannedStart { get; set; }
        public DateTime? PlannedFinish { get; set; }

        public bool Add()
        {
            string iSql = @"INSERT INTO JDE_Processes (ActionTypeId, PlaceId, CreatedBy, CreatedOn, TenantId, PlannedStart, PlannedFinish, LastStatus, LastStatusBy, LastStatusOn, IsActive, IsFrozen, IsCompleted, IsSuccessfull)
                            output INSERTED.ProcessId 
                            VALUES(@ActionTypeId, @PlaceId, @CreatedBy, @CreatedOn, @TenantId, @PlannedStart, @PlannedFinish, @LastStatus, @LastStatusBy, @LastStatusOn, @IsActive, @IsFrozen, @IsCompleted, @IsSuccessfull)";

            using (SqlCommand command = new SqlCommand(iSql, Settings.conn))
            {
                command.Parameters.AddWithValue("@ActionTypeId", 2);
                command.Parameters.AddWithValue("@PlaceId", Place.PlaceId);
                command.Parameters.AddWithValue("@CreatedBy", Settings.CurrentUser.UserId);
                command.Parameters.AddWithValue("@CreatedOn", DateTime.Now);
                command.Parameters.AddWithValue("@TenantId",1);
                command.Parameters.AddWithValue("@PlannedStart", PlannedStart);
                command.Parameters.AddWithValue("@PlannedFinish", PlannedFinish);
                command.Parameters.AddWithValue("@LastStatus", 1);
                command.Parameters.AddWithValue("@LastStatusBy", Settings.CurrentUser.UserId);
                command.Parameters.AddWithValue("@LastStatusOn", DateTime.Now);
                command.Parameters.AddWithValue("@IsActive", false);
                command.Parameters.AddWithValue("@IsFrozen", false);
                command.Parameters.AddWithValue("@IsCompleted", false);
                command.Parameters.AddWithValue("@IsSuccessfull", false);

                int result = -1;
                try
                {
                    result = (int)command.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Wystąpił błąd przy dodawaniu zgłoszenia dla zasobu {Place.Name}. Opis błędu: {ex.Message}");
                }

                if (result < 0)
                {
                    return false;
                }
                else
                {
                    ProcessId = result;
                    return true;
                }
            }
        }
    }
}
