using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using JdeScanExcelAddIn.Models;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace JdeScanExcelAddIn
{
    public partial class JdeScanRibbon
    {
        List<string> mUsers = new List<string>();
        List<string> mPlaces = new List<string>();
        List<string> aUsers = new List<string>();
        List<string> aPlaces = new List<string>();

        private void JdeScanRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //Static.Functions.ConfigNLog3();
        }

        private void btnJdeScanExport_Click(object sender, RibbonControlEventArgs e)
        {
            
            UsersKeeper uKeeper = new UsersKeeper();
            PlaceKeeper pKeeper = new PlaceKeeper();
            ActionKeeper aKeeper = new ActionKeeper();
            RecordKeeper rKeeper = new RecordKeeper();
            ActionTypeKeeper atKeeper = new ActionTypeKeeper();

            mPlaces.Clear();
            aPlaces.Clear();
            mUsers.Clear();
            aUsers.Clear();
            
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;
            Range UsedRange = sht.UsedRange;
            bool found = false;
            int cUser = 0;
            int cTime = 0;
            int cAction = 0;
            int cPlace = 0;
            int cType = 0;

            try
            {
                for (int i = 1; i <= UsedRange.Columns.Count; i++)
                {
                    if (cUser != 0 && cTime != 0 && cAction != 0 && cPlace != 0 && cType != 0)
                    {
                        found = true;
                        break;
                    }
                    else
                    {

                        string aCell = ((Range)UsedRange.Cells[1, i]).Value;

                        if (aCell == "Nazwa maszyny")
                        {
                            cPlace = i;
                        }
                        else if (aCell == "Czynność")
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
                    atKeeper.Reload();

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
                        if (!string.IsNullOrEmpty(names) && names != "Nazwisko" && !string.IsNullOrEmpty(act) && act != "Czynność" && !string.IsNullOrEmpty(pl) && pl != "Nazwa maszyny")
                        {
                            //Process only rows having any user assigned
                            //Don't do anything if the row doesn't contain Place and action
                            //They are indispensable

                            var nms = Regex.Split(names, ",");
                            if (nms.Count() == 1)
                            {
                                //Only 1 user? Or maybe those bustards are divided with "/" ?!
                                nms = Regex.Split(names, "/");
                                if (nms.Count() == 1)
                                {
                                    //Only 1 user? maybe backslash ("\") ?!
                                    if (names.Contains(@"\"))
                                    {
                                        nms = names.Split('\\');
                                    }
                                }
                            }
                            record.UsersAssigned = nms.Count();
                            foreach (string n in nms)
                            {
                                if (uKeeper.Items.Where(i => i.FullName == n.Trim()).Any())
                                {
                                    if (uKeeper.Items.Where(i => i.FullName == n.Trim() && i.IsArchived==true).Any() && !uKeeper.Items.Where(i => i.FullName == n.Trim() && i.IsArchived != true).Any())
                                    {
                                        //check if there is archived user like that but also check if there is active user like this
                                        //add it to archived list
                                        aUsers.Add(n.Trim());
                                    }
                                    else
                                    {
                                        //Keep User with id in record object
                                        record.Users.Add(new User { UserId = uKeeper.Items.Where(i => i.FullName == n.Trim() && i.IsArchived!=true).FirstOrDefault().UserId, FullName = n.Trim() });
                                    }     
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

                            if (pKeeper.Items.Where(i => i.Name.Trim() == pl).Any())
                            {
                                //Keep place with id in record object
                                if(pKeeper.Items.Where(i=>i.Name.Trim() == pl && i.IsArchived == true).Any() && !pKeeper.Items.Where(i => i.Name.Trim() == pl && i.IsArchived != true).Any())
                                {
                                    //check if there is archived place like that but also check if there is active place like this
                                    //add it to archived list
                                    aPlaces.Add(pl);
                                }
                                else
                                {
                                    record.Place = pKeeper.Items.Where(i => i.Name.Trim() == pl && i.IsArchived!=true).FirstOrDefault();
                                }
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

                    if (rKeeper.Items.Where(r => r.IsValid == false).Any())
                    {
                        IsValid = false;
                        //if at least one record contains invalid data, notify him
                        //and check if he would like to add valid records or correct it first
                        string mess = "Nie wszystkie rekordy można dodać. Jeśli w wierszu nazwa maszyny (lub nazwisko pracownika) nie występuje w programie, takiego wiersza nie można dodać. Aby dodać niepoprawne wiersze (zaznaczone na czerwono), należy najpierw utworzyć odpowiednią maszynę/użytkownika w programie, lub skorygować dane w pliku.";
                        if (mUsers.Any())
                            mess += Environment.NewLine + Environment.NewLine + "Brakujący użytkownicy: " + string.Join(", ", mUsers);
                        if (mPlaces.Any())
                            mess += Environment.NewLine + Environment.NewLine + "Brakujące maszyny: " + string.Join(", ", mPlaces);
                        if (aPlaces.Any() || aUsers.Any())
                        {
                            mess += Environment.NewLine + Environment.NewLine + "Na liście znajdują się zarchiwizowane maszyny lub użytkownicy.";
                            if(aPlaces.Any())
                                mess += Environment.NewLine + "Zarchiwizowane maszyny: " + string.Join(", ", aPlaces.Distinct());
                            if (aUsers.Any())
                                mess += Environment.NewLine + "Zarchiwizowani użytkownicy: " + string.Join(", ", aUsers.Distinct());
                        }
                            
                        mess += Environment.NewLine + Environment.NewLine + "Chcesz importować teraz poprawne wiersze (ZIELONE)?";
                        DialogResult res = MessageBox.Show(mess, "Niepoprawne dane", MessageBoxButtons.YesNo);
                        if (res == DialogResult.Yes)
                        {
                            IsValid = true;
                        }
                    }

                    if (IsValid)
                    {
                        //get week number
                        frmPeriod FrmPeriod = new frmPeriod();
                        DialogResult res = FrmPeriod.ShowDialog();
                        if (res == DialogResult.OK)
                        {
                            //All set, let's import the motherfucker
                            int w = (int)FrmPeriod.week;
                            int y = (int)FrmPeriod.year;
                            rKeeper.PlannedStart = (DateTime)FrmPeriod.startDate;
                            rKeeper.PlannedFinish = ((DateTime)FrmPeriod.startDate).AddHours(160);
                            frmLogin FrmLogin = new frmLogin(uKeeper);
                            DialogResult result = FrmLogin.ShowDialog();
                            if (result == DialogResult.OK)
                            {
                                Globals.ThisAddIn.Application.StatusBar = $"Importuje dane dla tygodnia {w}/{y}..";
                                int added = rKeeper.ImportAll();
                                Globals.ThisAddIn.Application.StatusBar = "";
                                MessageBox.Show("Import zakończony powodzeniem!", "Powodzenie");
                            }
                            else
                            {
                                //The user aborted the form and we don't have UserId of loagged user to upload to
                                MessageBox.Show("Akcja przerwana przez użytkownika. Żadne dane nie zostały zaimportowane..", "Import przerwany");
                            }
                        }
                        else
                        {
                            //The user aborted the form and we don't have week/year to upload to
                            MessageBox.Show("Akcja przerwana przez użytkownika. Żadne dane nie zostały zaimportowane..", "Import przerwany");
                        }

                    }
                }
            }catch(Exception ex)
            {
                MessageBox.Show($"Wystąpił błąd podczas analizowania pliku: {ex.Message}", "Import przerwany");
            }
        }

        private async void btnPlacePriority_Click(object sender, RibbonControlEventArgs e)
        {
            DateTime start = DateTime.Now;
            PlaceKeeper pKeeper = new PlaceKeeper();
            PlaceKeeper lKeeper = new PlaceKeeper();
            pKeeper.Reload();
            lKeeper = await GetRanking();
            int num = 0;
            foreach (Place p in pKeeper.Items)
            {
                if (lKeeper.Items.Where(i => i.Name.Trim() == p.Name.Trim()).Any())
                {
                    //if there's a match between file & db, update data from db
                    p.Priority = lKeeper.Items.Where(i => i.Name.Trim() == p.Name.Trim()).FirstOrDefault().Priority;
                    p.IsUpdated = true;
                    num++;
                }
            }
            
            DialogResult res = MessageBox.Show($"Znaleziono {num} pasujących maszyn. Czy chcesz zaktualizować priorytet maszyny danymi z pliku?", "Potwierdź", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(res == DialogResult.Yes)
            {
                frmLooper Looper = new frmLooper();
                Looper.Show();
                await pKeeper.UpdatePriority();
                int updated = pKeeper.Items.Count(i => i.IsUpdated == false);
                if (updated > 0)
                {
                    DateTime end = DateTime.Now;
                    MessageBox.Show($"Zaktualizowano {updated} maszyn w {(end-start).TotalSeconds} sekund!", "Powodzenie", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Niezaktualizowano żadnej maszyny, ponieważ żadna z maszyn w pliku nie została odnaleziona w bazie. Możliwe, że trzeba będzie założyć brakujące maszyny..", "Brak nowych danych", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                Looper.Hide();
            }
            

            
        }

        private async Task<PlaceKeeper> GetRanking()
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = null;
            PlaceKeeper lKeeper = new PlaceKeeper();
            bool found = false;
            int cPlace = 0;
            int cPriority = 0;

            try
            {
                try
                {
                    sht = wb.Sheets["Lista maszyn"];

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Nie znaleziono akrusza \"Lista maszyn\"..", "Niepowodzenie", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw;
                }
                Range UsedRange = sht.UsedRange;
                for (int i = 1; i <= UsedRange.Columns.Count; i++)
                {
                    if (cPlace > 0 && cPriority > 0)
                    {
                        found = true;
                        break;
                    }
                    else
                    {

                        string aCell = ((Range)UsedRange.Cells[2, i]).Value;

                        if (aCell == "Nazwa")
                        {
                            cPlace = i;
                        }
                        else if (aCell == "Ranking")
                        {
                            cPriority = i;
                        }
                    }
                }

                if (!found)
                {
                    //Not all variables have been found
                    MessageBox.Show("Nie udało się znaleźć wszystkich potrzebnych kolumn (Nazwa, Ranking). Upewnij się, że raport zawiera wszystkie kolumny z nagłówkami w pierwszym wierszu.", "Niepowodzenie", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    foreach (Range Row in UsedRange.Rows)
                    {
                        string pl = null;
                        if (((Range)UsedRange[Row.Row, cPlace]).Value2 != null)
                            pl = ((Range)UsedRange[Row.Row, cPlace]).Value.ToString().Trim();
                        string pr = null;
                        if (((Range)UsedRange[Row.Row, cPriority]).Value2 != null)
                            pr = ((Range)UsedRange[Row.Row, cPriority]).Value.ToString().Trim();
                        
                        if(pl!=null && pr != null)
                        {
                            //only if both place & priority has value
                            if(!lKeeper.Items.Where(i=>i.Name.Trim() == pl).Any())
                            {
                                //we don't have this place yet
                                Place Place = new Place();
                                Place.Name = pl;
                                Place.Priority = pr;
                                lKeeper.Items.Add(Place);
                            }
                        }
                    }
                }

                
            }
            catch (Exception ex)
            {

                MessageBox.Show("Aktualizacja maszyn zakończona niepowodzeniem, ponieważ nie udało się odczytać rankingu.","Niepowodzenie",MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return lKeeper;
        }
    }
}
