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
    public partial class frmPeriod : Form
    {
        public int? week { get; set; }
        public int? year { get; set; }
        public DateTime? startDate { get; set; }

        public frmPeriod()
        {
            InitializeComponent();
        }

        private void txtWeek_TextChanged(object sender, EventArgs e)
        {
            Calculate();
        }

        private void Calculate()
        {
            int w=0;
            int y=0;

            if(!string.IsNullOrEmpty(txtWeek.Text))
            {
                bool num = int.TryParse(txtWeek.Text, out w);
            }

            if (!string.IsNullOrEmpty(txtYear.Text))
            {
                bool num = int.TryParse(txtYear.Text, out y);
            }

            if (w > 0)
                week = w;
            if (y > 0)
                year = y;
            if(year > 0 && week > 0)
            {
                //do the calculation
                startDate = Static.Functions.FirstDateOfWeek((int)year, (int)week);
                startDate = ((DateTime)startDate).AddHours(-2);
                txtDates.Text = startDate.ToString() + " - " + ((DateTime)startDate).AddHours(160).ToString();
            }
            else
            {
                txtDates.Text = "";
            }

        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            Calculate();
        }
    }
}
