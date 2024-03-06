using Microsoft.Data.SqlClient; 
using System.Data; 

namespace Prescient
{
    public class SqlBulkInserter
    {
        private string _connectionString; // Field to store the connection string.

        public SqlBulkInserter(string connectionString)
        {
            _connectionString = connectionString; // Constructor to initialize the connection string.
        }

        // Method to perform bulk insertion into SQL Server table.
        public void BulkInsert(DataTable dataTable)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString)) // Creating a SQL connection object.
                {
                    connection.Open(); // Opening the database connection.

                    // Creating a SQL bulk copy instance.
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                    {
                        // Setting the destination table name for bulk insertion.
                        bulkCopy.DestinationTableName = "DailyMTM";

                        // Setting up the column mappings between source and destination tables.
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                        }

                        // Performing the bulk copy operation.
                        bulkCopy.WriteToServer(dataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error performing bulk insertion: {ex.Message}"); // Handling and displaying any exceptions that occur during bulk insertion.
            }
        }
    }
}
