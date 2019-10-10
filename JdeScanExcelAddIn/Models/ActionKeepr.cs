using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JdeScanExcelAddIn.Models
{
    public class ActionKeepr : Keeper<Action>
    {
        public void SaveAll()
        {
            foreach (Action item in Items)
            {
                item.Add(conn);
            }
        }
    }
}
