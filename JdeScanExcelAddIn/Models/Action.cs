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
        public int? GivenTime { get; set; }
        public string Type { get; set; }

        public bool Add(SqlConnection connection)
        {
            string iSql = @"INSERT INTO JDE_Actions (Name, CreatedBy, CreatedOn, TenantId, GivenTime, Type) 
                            VALUES(@Name, @CreatedBy, @CreatedOn, @TenantId, @GivenTime, @Type)";
            using (SqlCommand command = new SqlCommand(iSql,connection))
            {
                command.Parameters.AddWithValue("@Name", Name);
                command.Parameters.AddWithValue("@GivenTime", GivenTime);
                command.Parameters.AddWithValue("@Type", Type);
                command.Parameters.AddWithValue("@CreatedBy", 1);
                command.Parameters.AddWithValue("@CreatedOn", DateTime.Now);
                command.Parameters.AddWithValue("@TenantId", 1);

                if(connection.State != System.Data.ConnectionState.Open)
                    connection.Open();

                int result = command.ExecuteNonQuery();
                if (result < 0)
                {
                    MessageBox.Show($"Wystąpił błąd przy dodawaniu czynności {Name} do bazy");
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
