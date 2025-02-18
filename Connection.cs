using System;
using System.Data.SqlClient;

namespace CARDIOLOGY
{
    public class Connection
    {
        public readonly string connectionString;

        public Connection()
        {
            //connectionString = Properties.Settings.Default.ConnectServer;
        }

        public SqlConnection OpenConnection()
        {
            SqlConnection connection = new SqlConnection(connectionString);
            try
            {
                connection.Open();
                return connection;
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to open the database connection: " + ex.Message);
            }
        }

        public void CloseConnection(SqlConnection connection)
        {
            if (connection != null && connection.State == System.Data.ConnectionState.Open)
            {
                try
                {
                    connection.Close();
                }
                catch (Exception ex)
                {
                    throw new Exception("Failed to close the database connection: " + ex.Message);
                }
            }
        }
    }
}
