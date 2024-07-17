using System.Data.SqlClient;
using System.Data;

namespace B1PostingDataText.Connection
{
    public class DapperConnection
        {
            protected static IDbConnection SQLOpenConnection()
            {
                string connectionstring = Properties.Settings.Default.SQLConnections; ;
                IDbConnection dbConnection = new SqlConnection(connectionstring);
                dbConnection.Open();
                return dbConnection;
            }

        }
}
