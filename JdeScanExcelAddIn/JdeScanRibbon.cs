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
        List<string> mUsers = new List<string>();


        private void JdeScanRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnJdeScanExport_Click(object sender, RibbonControlEventArgs e)
        {
            frmPeriod FrmPeriod = new frmPeriod();
            FrmPeriod.ShowDialog();
            UsersKeeper uKeeper = new UsersKeeper();
            
            ActionKeeper aKeeper = new ActionKeeper();
            
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;
            Range UsedRange = sht.UsedRange;
            bool found = false;
            int cUser = 0;
            int cTime = 0;
            int cAction = 0;
            int cPlace = 0;
            int cType = 0;

            for (int i = 1; i <= UsedRange.Columns.Count; i++)
            {
                if (cUser!=0 && cTime!=0 && cAction!=0 && cPlace!=0 && cType != 0)
                {
                    found = true;
                    break;
                }
                else
                {
                    
                    string aCell = ((Range)UsedRange.Cells[1, i]).Value;
                    
                    if (aCell== "Nazwa maszyny")
                    {
                        cPlace = i;
                    }else if(aCell == "Czynność")
                    {
                        cAction = i;
                    }
                    else if (aCell == "S/R")
                    {
                        cType = i;
                    }
                    else if (aCell == "czas")
                    {
                        cTime = i;
                    }
                    else if (aCell == "Nazwisko")
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
                uKeeper.Reload();
                aKeeper.Reload();

                foreach (Range Row in UsedRange.Rows)
                {
                    //Go and add missing actions to db
                    string act = ((Range)UsedRange[Row.Row, cAction]).Value;
                    if (!string.IsNullOrEmpty(act) && act != "Czynność")
                    {
                        if (!aKeeper.Items.Where(i => i.Name == act).Any())
                        {
                            //it doesn't exist, let's add it
                            Models.Action a = new Models.Action();
                            a.Name = act;
                            int min = 0;
                            bool passed = int.TryParse(((Range)UsedRange[Row.Row, cTime]).Value, out min);
                            a.GivenTime = min;
                            a.Type = ((string)((Range)UsedRange[Row.Row, cTime]).Value).Trim();
                            if (a.Add())
                            {
                                aKeeper.Items.Add(a);
                            }
                        }
                    }
                }

                foreach (Range Row in UsedRange.Rows)
                {
                    //get Users
                    string names = ((Range)UsedRange[Row.Row, cUser]).Value;
                    if (!string.IsNullOrEmpty(names) && names!="Nazwisko")
                    {
                        var nms = Regex.Split(names, ",");
                        if (nms.Count() == 1)
                        {
                            //Only 1 user? Or maybe those bustards are divided with "/" ?!
                            nms = Regex.Split(names, "/");
                            if(nms.Count() == 1)
                            {
                                //Only 1 user? maybe backslash ("\") ?!
                                nms = Regex.Split(names, "\\");
                            }
                        }
                        foreach(string n in nms)
                        {
                            if(uKeeper.Items.Where(i=>i.FullName == n.Trim()).Any())
                            {
                                Users.Add(new User { UserId = uKeeper.Items.Where(i=>i.FullName ==n.Trim()).FirstOrDefault().UserId, FullName = n.Trim() });
                            }
                            else
                            {
                                if (!mUsers.Where(i => i == n.Trim()).Any())
                                {
                                    mUsers.Add(n.Trim());
                                }
                                
                            }
                            
                        }
                    }

                    //get Actions
                    string act = ((Range)UsedRange[Row.Row, cAction]).Value;
                    if(!string.IsNullOrEmpty(act) && act != "Czynność")
                    {

                    }

                }

                if (mUsers.Any())
                {
                    MessageBox.Show($"Na liście użytkowników programu brakuje {mUsers.Count} pozycj, który znajdują się w pliku w kolumnie Nazwisko. Brakujące pozycje: {string.Join(", ", mUsers)}. Przerywam export.");
                }
            }

        }
    }
}
