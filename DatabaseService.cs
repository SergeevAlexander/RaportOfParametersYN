using Oracle.ManagedDataAccess.Client;
using System.Configuration;
using System.Collections.Generic;
using System.Data;
using System;

namespace ExcelEf.Services
{
    internal class DatabaseService
    {// Запрос для того, чтобы коннектиться к базе
        internal List<Dictionary<string, object>> ExecuteQuery(string sql)
        {
            var results = new List<Dictionary<string,object>>();

            using (var connection = new OracleConnection(DatabaseConfig.ConnectionString))
            {
                connection.Open();
                using (var command = new OracleCommand(sql, connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var row = new Dictionary<string, object>();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                string columnName = reader.GetName(i);
                                row[columnName] = (reader.IsDBNull(i)) ? null : reader.GetValue(i);
                            }
                            results.Add(row);
                        }
                    }
                }
            }
            return results;
        }
    }
}
