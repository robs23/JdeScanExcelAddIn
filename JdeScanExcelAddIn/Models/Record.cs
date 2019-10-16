using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JdeScanExcelAddIn.Models
{
    public class Record
    {
        public int RowNumber { get; set; }
        public bool IsValid
        {
            get
            {
                if(Users.Count == UsersAssigned && Place != null)
                {
                    return true;
                }
                return false;
            }
        }

        public int UsersAssigned { get; set; }
        public List<User> Users { get; set; }
        public Place Place { get; set; }
        public Action Action { get; set; }

        public Record()
        {
            Users = new List<User>();
        }
    
    }
}
