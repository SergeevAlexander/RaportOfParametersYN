using System.Configuration;

namespace ExcelEf
{
    internal static class DatabaseConfig
    {
        public static string ConnectionString
        {
          get { return ConfigurationManager.ConnectionStrings["OracleDB"].ConnectionString; }
        }
        public static string ConnectionStringCDS
        {
            get { return ConfigurationManager.ConnectionStrings["OracleDBCDS"].ConnectionString; }
        }
    }
}
