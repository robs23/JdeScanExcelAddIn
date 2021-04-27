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
                if(Place != null)
                {
                    if (ActionType.RequireUsersAssignment!=true)
                    {
                        if (UsersAssigned == 0 || Users.Count == UsersAssigned)
                        {
                            return true;
                        }                    
                    }
                    else
                    {
                        if(Users.Count == UsersAssigned && Users.Count > 0)
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
        }

        public int UsersAssigned { get; set; }
        public List<User> Users { get; set; }
        public Place Place { get; set; }
        public Action Action { get; set; }
        public Process Process { get; set; }
        public ActionType ActionType { get; set; }

        public Record()
        {
            Users = new List<User>();
        }
    
    }
}
