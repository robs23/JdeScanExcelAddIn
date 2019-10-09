using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JdeScanExcelAddIn.Models
{
    public class UsersKeeper : Keeper<User>
    {
        protected override string ObjectName => "User";

        protected override string PluralizedObjectName => "Users";

    }
}
