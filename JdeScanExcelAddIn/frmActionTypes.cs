using JdeScanExcelAddIn.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JdeScanExcelAddIn
{
    public partial class frmActionTypes : Form
    {
        public ActionType chosen { get; set; } = null;
        public ActionTypeKeeper Keeper { get; set; }
        public frmActionTypes(ActionTypeKeeper ActionTypeKeeper)
        {
            InitializeComponent();
            Keeper = ActionTypeKeeper;
            cmbActionType.DataSource = Keeper.Items;
            cmbActionType.DisplayMember = "Name";
            cmbActionType.ValueMember = "ActionTypeId";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            chosen = Keeper.Items.FirstOrDefault(a=>a.ActionTypeId == (int)cmbActionType.SelectedValue);
            DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
