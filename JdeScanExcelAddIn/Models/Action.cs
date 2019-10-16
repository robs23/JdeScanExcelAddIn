using JdeScanExcelAddIn.Static;
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
    public class Action : Entity<Action>
    {
        [DisplayName("ID")]
        public int ActionId { get; set; }
        public override int Id
        {
            set => value = ActionId;
            get => ActionId;
        }
        public string Name { get; set; }
        public Nullable<int> GivenTime { get; set; } = null;
        public string Type { get; set; }

        public bool Add()
        {
            string iSql = @"INSERT INTO JDE_Actions (Name, CreatedBy, CreatedOn, TenantId, GivenTime, Type)
                            output INSERTED.ActionID 
                            VALUES(@Name, @CreatedBy, @CreatedOn, @TenantId, @GivenTime, @Type)";

            using (SqlCommand command = new SqlCommand(iSql, Settings.conn))
            {
                command.Parameters.AddWithValue("@Name", Name);
                command.Parameters.AddWithNullableValue("@GivenTime", GivenTime);
                command.Parameters.AddWithNullableValue("@Type", Type);
                command.Parameters.AddWithValue("@CreatedBy", 1);
                command.Parameters.AddWithValue("@CreatedOn", DateTime.Now);
                command.Parameters.AddWithValue("@TenantId", 1);

                int result = -1;
                try
                {
                    result = (int)command.ExecuteScalar();
                }
                catch(Exception ex)
                {
                    MessageBox.Show($"Wystąpił błąd przy dodawaniu czynności {Name} do bazy. Opis błędu: {ex.Message}");
                }
                
                if (result < 0)
                {
                    return false;
                }
                else
                {
                    ActionId = result;
                    return true;
                }
            }
        }
    }
}
