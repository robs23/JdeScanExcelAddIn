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
        List<string> mUsers = new List<string>();
        List<string> mPlaces = new List<string>();

        private void JdeScanRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnJdeScanExport_Click(object sender, RibbonControlEventArgs e)
        {
            
            UsersKeeper uKeeper = new UsersKeeper();
            PlaceKeeper pKeeper = new PlaceKeeper();
            ActionKeeper aKeeper = new ActionKeeper();
            RecordKeeper rKeeper = new RecordKeeper();
            
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
                pKeeper.Reload();

                foreach (Range Row in UsedRange.Rows)
                {
                    Record record = new Record();
                    record.RowNumber = Row.Row;
                    //get Users
                    string names = null;
                    if (((Range)UsedRange[Row.Row, cUser]).Value2 != null)
                        names = ((Range)UsedRange[Row.Row, cUser]).Value;
                    string act = null;
                    if (((Range)UsedRange[Row.Row, cAction]).Value2 != null)
                        act = ((Range)UsedRange[Row.Row, cAction]).Value.ToString().Trim();
                    string pl = null;
                    if (((Range)UsedRange[Row.Row, cPlace]).Value2 != null)
                        pl = ((Range)UsedRange[Row.Row, cPlace]).Value.ToString().Trim();
                    if (!string.IsNullOrEmpty(names) && names!="Nazwisko" && !string.IsNullOrEmpty(act) && act != "Czynność" && !string.IsNullOrEmpty(pl) && pl != "Nazwa maszyny")
                    {
                        //Process only rows having any user assigned
                        //Don't do anything if the row doesn't contain Place and action
                        //They are indispensable

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
                        record.UsersAssigned = nms.Count();
                        foreach(string n in nms)
                        {
                            if(uKeeper.Items.Where(i=>i.FullName == n.Trim()).Any())
                            {
                                //Keep User with id in record object
                                record.Users.Add(new User { UserId = uKeeper.Items.Where(i=>i.FullName ==n.Trim()).FirstOrDefault().UserId, FullName = n.Trim() });
                            }
                            else
                            {
                                if (!mUsers.Where(i => i == n.Trim()).Any())
                                {
                                    //add it to missing list
                                    mUsers.Add(n.Trim());
                                }
                                
                            }
                            
                        }

                        if(pKeeper.Items.Where(i=>i.Name.Trim() == pl).Any())
                        {
                            //Keep place with id in record object
                            record.Place = pKeeper.Items.Where(i => i.Name.Trim() == pl).FirstOrDefault();
                        }
                        else
                        {
                            if (!mPlaces.Where(i => i == pl).Any())
                            {
                                //add it to missing list
                                mPlaces.Add(pl);
                            }
                        }


                        //get Actions
                        Models.Action a = new Models.Action();

                        if (!aKeeper.Items.Where(i => i.Name.Trim().Equals(act, StringComparison.OrdinalIgnoreCase)).Any())
                        {
                            //Go and add missing actions to db
                            //it doesn't exist, let's add it
                            a.Name = act;
                            bool passed = false;
                            int min = 0;
                            string tg = null;
                            try
                            {
                                if (((Range)UsedRange[Row.Row, cTime]).Value2 != null)
                                {
                                    tg = ((Range)UsedRange[Row.Row, cTime]).Value.ToString();
                                    passed = int.TryParse(tg, out min);
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                            if (passed)
                                a.GivenTime = min;

                            if (((Range)UsedRange[Row.Row, cType]).Value2 != null)
                                a.Type = ((Range)UsedRange[Row.Row, cType]).Value.ToString().Trim();

                        }
                        else
                        {
                            a = aKeeper.Items.Where(i => i.Name.Trim().Equals(act, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
                        }

                        record.Action = a;
                        rKeeper.Items.Add(record);
                    }

                    if (!record.IsValid)
                    {
                        // mark it
                        Row.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    else
                    {
                        Row.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                    }
                    
                }


                bool IsValid = true;

                if (rKeeper.Items.Where(r=>r.IsValid==false).Any())
                {
                    IsValid = false;
                    //if at least one record contains invalid data, notify him
                    //and check if he would like to add valid records or correct it first
                    string mess = "Nie wszystkie rekordy można dodać. Jeśli w wierszu nazwa maszyny (lub nazwisko pracownika) nie występuje w programie, takiego wiersza nie można dodać. Aby dodać niepoprawne wiersze (zaznaczone na czerwono), należy najpierw utworzyć odpowiednią maszynę/użytkownika w programie, lub skorygować dane w pliku.";
                    if(mUsers.Any())
                        mess += Environment.NewLine + Environment.NewLine + "Brakujący użytkownicy: " + string.Join(", ", mUsers);
                    if(mPlaces.Any())
                        mess += Environment.NewLine + Environment.NewLine + "Brakujące maszyny: " + string.Join(", ", mPlaces);
                    mess += Environment.NewLine + Environment.NewLine + "Chcesz importować teraz poprawne wiersze (NIE czerwone)?";
                    DialogResult res = MessageBox.Show(mess, "Niepoprawne dane", MessageBoxButtons.YesNo);
                    if(res == DialogResult.Yes)
                    {
                        IsValid = true;
                    }
                }

                if (IsValid)
                {
                    //get week number
                    frmPeriod FrmPeriod = new frmPeriod();
                    DialogResult res = FrmPeriod.ShowDialog();
                    if(res == DialogResult.OK)
                    {
                        //All set, let's import the motherfucker
                        int w = (int)FrmPeriod.week;
                        int y = (int)FrmPeriod.year;
                        DateTime startDate = (DateTime)FrmPeriod.startDate;
                        Globals.ThisAddIn.Application.StatusBar = $"Importuje dane dla tygodnia {w}/{y}..";
                        Import(rKeeper);
                        //MessageBox.Show($"Importuje dane dla tygodnia {w}/{y}.", "Przygotowany do importu");
                    }
                    else
                    {
                        //The user aborted the form and we don't have week/year to upload to
                        MessageBox.Show("Akcja przerwana przez użytkownika. Żadne dane nie zostały zaimportowane..", "Import przerwany");
                    }

                }
            }

        }

        private bool Import(RecordKeeper rKeeper)
        {
            bool status = false;
            int rCount = 0;
            List<string> failedActions = new List<string>();

            //add missing actions first
            try
            {
                foreach (Record r in rKeeper.Items.Where(i => i.Action.ActionId == 0))
                {
                    if (r.Action.Add())
                    {
                        rCount++;
                    }
                    else
                    {
                        //
                        failedActions.Add(r.Action.Name);
                    }
                }
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            

            //add missing PlaceActions

            int importedPlaceActions = rKeeper.ImportPlaceActions(); 

            return status;
        }
    }
}
