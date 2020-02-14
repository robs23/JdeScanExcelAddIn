using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JdeScanExcelAddIn.Static
{
    public static class Extensions
    {
        public static SqlParameter AddWithNullableValue(this SqlParameterCollection collection, string parameterName, object value)
        {
            if (value == null)
                return collection.AddWithValue(parameterName, DBNull.Value);
            else
                return collection.AddWithValue(parameterName, value);
        }

        public static T GetValueOrDefault<T>(this SqlDataReader dataReader, string columnName)
        {
            //checks if cell contain null and if so, converts value to null
            return !dataReader.IsDBNull(dataReader.GetOrdinal(columnName)) ? (T)dataReader.GetValue(dataReader.GetOrdinal(columnName)) : default(T);
        }
    }
}
