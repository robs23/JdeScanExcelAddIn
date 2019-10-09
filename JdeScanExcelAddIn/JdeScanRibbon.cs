using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;

namespace JdeScanExcelAddIn
{
    public partial class JdeScanRibbon
    {
        private void JdeScanRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnJdeScanExport_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;
            MessageBox.Show(sht.Name);
        }
    }
}
