using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JdeScanExcelAddIn.Models
{
    public class RecordKeeper : Keeper<Record>
    {
        public List<Process> Processes { get; set; }
        public List<Action> Actions { get; set; }
        public int RowsAdded { get; set; } = 0;
        public DateTime PlannedStart { get; set; }
        public DateTime PlannedFinish { get; set; }

        public RecordKeeper()
        {
            Processes = new List<Process>();
            Actions = new List<Action>();
        }

        public int ImportAll()
        {
            RowsAdded += ImportActions();
            RowsAdded += ImportPlaceActions();
            RowsAdded += ImportProcesses();
            RowsAdded += ImportProcessActions();
            RowsAdded += ImportProcessAssigns();

            return RowsAdded;
        }


        public int ImportActions()
        {
            int rCount = 0;
            List<string> failedActions = new List<string>();

            //add missing actions first
            try
            {
                foreach (Record r in Items.Where(i => i.Action.ActionId == 0))
                {
                    if(!Actions.Where(i=>i.Name.ToLower().Trim() == r.Action.Name.ToLower().Trim()).Any())
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
                        Actions.Add(r.Action);
                    }
                    else
                    {
                        r.Action.ActionId = Actions.Where(i => i.Name.ToLower().Trim() == r.Action.Name.ToLower().Trim()).FirstOrDefault().ActionId;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return rCount;

        }

        public void DeleteExistent()
        {
            bool existent = false;

            string vSql = "SELECT ProcessId FROM JDE_Processes WHERE PlannedStart = @PlannedStart AND PlannedFinish=@PlannedFinish AND IsActive=0 AND IsFrozen=0 AND IsCompleted=0";
            using (SqlCommand command = new SqlCommand(vSql, Settings.conn))
            {
                command.Parameters.AddWithValue("@PlannedStart", PlannedStart);
                command.Parameters.AddWithValue("@PlannedFinish", PlannedFinish);
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        existent = true;
                    }
                }
            }

            if (existent)
            {
                DialogResult userChoice = MessageBox.Show("Jeśli w wybranym okresie istnieją już planowane i nierozpoczęte zgłoszenia, czy chcesz je usunąć?", "Istniejące zgłoszenia", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (userChoice == DialogResult.Yes)
                {
                    //Delete planned processes' planned actions
                    string sql = @"DELETE 
                        FROM JDE_ProcessActions
                        WHERE ProcessId IN (" + vSql + ")";

                    using (SqlCommand command = new SqlCommand(sql, Settings.conn))
                    {
                        command.Parameters.AddWithValue("@PlannedStart", PlannedStart);
                        command.Parameters.AddWithValue("@PlannedFinish", PlannedFinish);
                        command.ExecuteNonQuery();
                    }

                    //Delete planned processes' planned actions
                    sql = @"DELETE 
                        FROM JDE_ProcessAssigns
                        WHERE ProcessId IN (" + vSql + ")";

                    using (SqlCommand command = new SqlCommand(sql, Settings.conn))
                    {
                        command.Parameters.AddWithValue("@PlannedStart", PlannedStart);
                        command.Parameters.AddWithValue("@PlannedFinish", PlannedFinish);
                        command.ExecuteNonQuery();
                    }

                    //Delete planned processes in the same period
                    sql = @"DELETE 
                        FROM JDE_Processes
                        WHERE PlannedStart = @PlannedStart AND PlannedFinish=@PlannedFinish AND IsActive=0 AND IsFrozen=0 AND IsCompleted=0";

                    using (SqlCommand command = new SqlCommand(sql, Settings.conn))
                    {
                        command.Parameters.AddWithValue("@PlannedStart", PlannedStart);
                        command.Parameters.AddWithValue("@PlannedFinish", PlannedFinish);
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        public int ImportProcesses()
        {
            DeleteExistent();

            int res = 0;

            foreach (Record r in Items.Where(i => i.IsValid==true))
            {
                if (!Processes.Where(i => i.Place.PlaceId == r.Place.PlaceId).Any())
                {
                    //we don't have this Place yet
                    Process p = new Process { Place = r.Place, PlannedStart = PlannedStart, PlannedFinish = PlannedFinish };
                    if (p.Add())
                    {
                        res++;
                    }

                    Processes.Add(p);
                    r.Process = p;
                }
                else
                {
                    r.Process = Processes.Where(i => i.Place.PlaceId == r.Place.PlaceId).FirstOrDefault();
                }
            }

            return res;

        }

        public int ImportProcessActions()
        {
            int res = -1;
            string cSql = "CREATE TABLE #ProcessActions(ProcessId int, ActionId int)";
            List<string> rStr = new List<string>(); //collection of records formatted for batch upload eg (1,2),(4,5),... Each item contains 1000 records max (sql server requirement)
            string cStr = ""; //current item
            int counter = 0;

            using (SqlCommand command = new SqlCommand(cSql, Settings.conn))
            {
                foreach (Record r in Items)
                {
                    //prepare insert string
                    counter++;
                    if (r.IsValid)
                    {
                        if (counter % 1000 == 0)
                        {
                            //we've just hit 1000 items

                            rStr.Add(cStr);
                            cStr = "";
                        }
                        cStr += $"({r.Process.ProcessId},{r.Action.ActionId}),";
                    }
                }
                //non-full item set must be added here... otherwise it won't be added
                if (!string.IsNullOrEmpty(cStr))
                    rStr.Add(cStr);

                if (rStr.Any())
                {

                    for (int i = 0; i < rStr.Count; i++)
                    {
                        rStr[i] = rStr[i].Substring(0, rStr[i].Length - 1); //drop the last ","
                    }

                }

                command.ExecuteNonQuery();

                if (rStr.Any())
                {
                    foreach (string s in rStr)
                    {
                        //do this for each 1000 items
                        string iSql = "INSERT INTO #ProcessActions(ProcessId, ActionId) VALUES " + s;
                        using (SqlCommand iCommand = new SqlCommand(iSql, Settings.conn))
                        {
                            iCommand.ExecuteNonQuery();
                        }
                    }

                    //once everything is uploaded to #ProcessActions, differentiate it with JDE_ProcessActions and add new items only
                    string sSql = $"SELECT DISTINCT ProcessId, ActionId, {Settings.CurrentUser.UserId} as CreatedBy, GETDATE() as CreatedOn, 1 as TenantId FROM #ProcessActions tpa WHERE NOT EXISTS (SELECT * FROM JDE_ProcessActions pa WHERE pa.ProcessId=tpa.ProcessId AND pa.ActionId=tpa.ActionId)";
                    string iiSql = "INSERT INTO JDE_ProcessActions (ProcessId, ActionId, CreatedBy, CreatedOn, TenantId) " + sSql;
                    using (SqlCommand iiCommand = new SqlCommand(iiSql, Settings.conn))
                    {
                        res = iiCommand.ExecuteNonQuery();
                    }
                }
            }

            return res;
        }

        public int ImportProcessAssigns()
        {
            int res = -1;
            string cSql = "CREATE TABLE #ProcessAssigns(ProcessId int, UserId int)";
            List<string> rStr = new List<string>(); //collection of records formatted for batch upload eg (1,2),(4,5),... Each item contains 1000 records max (sql server requirement)
            string cStr = ""; //current item
            int counter = 1;

            using (SqlCommand command = new SqlCommand(cSql, Settings.conn))
            {
                foreach (Record r in Items)
                {
                    //prepare insert string
                    
                    if (r.IsValid)
                    {
                        if (counter % 1000 == 0)
                        {
                            //we've just hit 1000 items

                            rStr.Add(cStr);
                            cStr = "";
                        }
                        foreach(User u in r.Users)
                        {
                            counter++;
                            cStr += $"({r.Process.ProcessId},{u.UserId}),";
                        }
                        
                    }
                }
                //non-full item set must be added here... otherwise it won't be added
                if (!string.IsNullOrEmpty(cStr))
                    rStr.Add(cStr);

                if (rStr.Any())
                {

                    for (int i = 0; i < rStr.Count; i++)
                    {
                        rStr[i] = rStr[i].Substring(0, rStr[i].Length - 1); //drop the last ","
                    }

                }

                command.ExecuteNonQuery();

                if (rStr.Any())
                {
                    foreach (string s in rStr)
                    {
                        //do this for each 1000 items
                        string iSql = "INSERT INTO #ProcessAssigns(ProcessId, UserId) VALUES " + s;
                        using (SqlCommand iCommand = new SqlCommand(iSql, Settings.conn))
                        {
                            iCommand.ExecuteNonQuery();
                        }
                    }

                    //once everything is uploaded to #ProcessAssigns, differentiate it with JDE_ProcessAssigns and add new items only
                    string sSql = $"SELECT DISTINCT ProcessId, UserId, {Settings.CurrentUser.UserId} as CreatedBy, GETDATE() as CreatedOn, 1 as TenantId FROM #ProcessAssigns tpa WHERE NOT EXISTS (SELECT * FROM JDE_ProcessAssigns pa WHERE pa.ProcessId=tpa.ProcessId AND pa.UserId=tpa.UserId)";
                    string iiSql = "INSERT INTO JDE_ProcessAssigns (ProcessId, UserId, CreatedBy, CreatedOn, TenantId) " + sSql;
                    using (SqlCommand iiCommand = new SqlCommand(iiSql, Settings.conn))
                    {
                        res = iiCommand.ExecuteNonQuery();
                    }
                }
            }

            return res;
        }

        public int ImportPlaceActions()
        {
            int res = -1;
            string cSql = "CREATE TABLE #PlaceActions(PlaceId int, ActionId int)";
            List<string> rStr = new List<string>() ; //collection of records formatted for batch upload eg (1,2),(4,5),... Each item contains 1000 records max (sql server requirement)
            string cStr = ""; //current item
            int counter = 0;

            using (SqlCommand command = new SqlCommand(cSql, Settings.conn))
            {
                foreach(Record r in Items)
                {
                    //prepare insert string
                    counter++;
                    if (r.IsValid) 
                    {
                        if(counter % 1000 == 0)
                        {
                            //we've just hit 1000 items
                            
                            rStr.Add(cStr);
                            cStr = "";
                        }
                        cStr += $"({r.Place.PlaceId},{r.Action.ActionId}),";
                    }
                }
                //non-full item set must be added here... otherwise it won't be added
                if(!string.IsNullOrEmpty(cStr))
                    rStr.Add(cStr);

                if (rStr.Any())
                {

                    for (int i = 0; i < rStr.Count; i++)
                    {
                        rStr[i] = rStr[i].Substring(0, rStr[i].Length - 1); //drop the last ","
                    }

                }

                command.ExecuteNonQuery();

                if (rStr.Any())
                {
                    foreach(string s in rStr)
                    {
                        //do this for each 1000 items
                        string iSql = "INSERT INTO #PlaceActions(PlaceId, ActionId) VALUES " + s;
                        using(SqlCommand iCommand = new SqlCommand(iSql, Settings.conn))
                        {
                            iCommand.ExecuteNonQuery();
                        }
                    }

                    //once everything is uploaded to #PlaceActions, differentiate it with JDE_PlaceActions and add new items only
                    string sSql = $"SELECT DISTINCT PlaceId, ActionId, {Settings.CurrentUser.UserId} as CreatedBy, GETDATE() as CreatedOn, 1 as TenantId FROM #PlaceActions tpa WHERE NOT EXISTS (SELECT * FROM JDE_PlaceActions pa WHERE pa.PlaceId=tpa.PlaceId AND pa.ActionId=tpa.ActionId)";
                    string iiSql = "INSERT INTO JDE_PlaceActions (PlaceId, ActionId, CreatedBy, CreatedOn, TenantId) " + sSql;
                    using(SqlCommand iiCommand = new SqlCommand(iiSql, Settings.conn))
                    {
                        res = iiCommand.ExecuteNonQuery();
                    }
                }
            }

            return res;
        }         
    
    }
}
