using Microsoft.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace Prescient
{
    class ExcelProcessor
    {
        private string _connectionString;

        public ExcelProcessor(string connectionString)
        {
            _connectionString = connectionString;
        }

        public void ProcessExcelFile(string filePath)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is on the first sheet

            // Get the used range of cells in the worksheet
            Excel.Range range = worksheet.UsedRange;

            // Connect to the SQL database
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                // Iterate through rows and extract contract details
                for (int row = 2; row <= range.Rows.Count; row++) // Assuming the first row is header
                {
                    string contract = (string)(range.Cells[row, 1] as Excel.Range).Value2;
                    string expiryDate = (string)(range.Cells[row, 3] as Excel.Range).Value2;

                    // Insert contract details into the database
                    InsertContractDetails(contract, expiryDate, connection);
                }
            }

            // Close Excel objects
            workbook.Close(false);
            excelApp.Quit();
        }

        private void InsertContractDetails(string contract, string expiryDate, SqlConnection connection)
        {
            string query = "INSERT INTO DailyMTM (Contract, ExpiryDate) VALUES (@Contract, @ExpiryDate)";
            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@Contract", contract);
                command.Parameters.AddWithValue("@ExpiryDate", expiryDate);
                // Add parameters for other fields as needed
                command.ExecuteNonQuery();
            }
        }
    }
}