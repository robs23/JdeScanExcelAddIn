using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JdeScanExcelAddIn.Models
{
    public class RecordKeeper : Keeper<Record>
    {

        public int ImportPlaceActions()
        {
            int res = -1;
            string cSql = "CREATE TABLE #PlaceActions(PlaceId int, ActionId int, TenantId int)";
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
                    string sSql = "SELECT DISTINCT PlaceId, ActionId, 40 as CreatedBy, GETDATE() as CreatedOn, 1 as TenantId FROM #PlaceActions tpa WHERE NOT EXISTS (SELECT * FROM JDE_PlaceActions pa WHERE pa.PlaceId=tpa.PlaceId AND pa.ActionId=tpa.ActionId)";
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
