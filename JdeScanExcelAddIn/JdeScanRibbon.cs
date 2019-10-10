using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using JdeScanExcelAddIn.Models;
using System.Text.RegularExpressions;

namespace JdeScanExcelAddIn
{
    public partial class JdeScanRibbon
    {
        List<User> Users = new List<User>();
        private void JdeScanRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnJdeScanExport_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;
            Range UsedRange = sht.UsedRange;
            bool found = false;
            int cUser = 0;
            int cTime = 0;
            int cAction = 0;
            int cPlace = 0;
            int cType = 0;

            for (int i = 1; i < UsedRange.Columns.Count; i++)
            {
                if (cUser!=0 && cTime!=0 && cAction!=0 && cPlace!=0 && cType != 0)
                {
                    found = true;
                    break;
                }
                else
                {
                    if(UsedRange.Cells[1,i]== "Nazwa maszyny")
                    {
                        cPlace = i;
                    }else if(UsedRange.Cells[1,i]== "Czynność")
                    {
                        cAction = i;
                    }
                    else if (UsedRange.Cells[1, i] == "S/R")
                    {
                        cType = i;
                    }
                    else if (UsedRange.Cells[1, i] == "czas")
                    {
                        cTime = i;
                    }
                    else if (UsedRange.Cells[1, i] == "Nazwisko")
                    {
                        cUser = i;
                    }
                }
            }

            if (!found)
            {
                //Not all variables have been found
                MessageBox.Show("Nie udało się znaleźć wszystkich potrzebnych kolumn (Nazwa maszyny, Czynność, S/R, czas, Nazwisko). Upewnij się, że raport zawiera wszystkie kolumny z nagłówkami w pierwszym wierszu.");
            }
            else
            {
                foreach (Range Row in UsedRange.Rows)
                {
                    string names = UsedRange[Row.Row, cUser];
                    if (!string.IsNullOrEmpty(names))
                    {
                        var nms = Regex.Split(names, ",");
                        foreach(string n in nms)
                        {
                            Users.Add(new User { FullName = n.Trim() });
                        }
                    }
                }
            }

        }
    }
}
