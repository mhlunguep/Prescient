using System.Runtime.InteropServices;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace Prescient
{
    public class ExcelProcessor
    {
        private string _connectionString; // Field to store the connection string.

        public ExcelProcessor(string connectionString)
        {
            _connectionString = connectionString; // Constructor to initialize the connection string.
        }

        // Method to process an Excel file.
        public void ProcessExcelFile(string filePath)
        {
            Excel.Application excelApp = new Excel.Application(); // Create an instance of Excel application.
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath); // Open the Excel workbook.
            Excel.Worksheet worksheet = workbook.Sheets[1]; // Get the first worksheet.

            try
            {
                ExcelDataExtractor extractor = new ExcelDataExtractor(); // Create an instance of ExcelDataExtractor.
                DataTable dataTable = extractor.ExtractData(worksheet); // Extract data from the Excel worksheet into a DataTable.

                SqlBulkInserter bulkInserter = new SqlBulkInserter(_connectionString); // Create an instance of SqlBulkInserter.
                bulkInserter.BulkInsert(dataTable); // Bulk insert the extracted data into SQL Server.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing Excel file: {ex.Message}"); // Handle and display any exceptions that occur during processing.
            }
            finally
            {
                workbook.Close(false); // Close the workbook without saving changes.
                excelApp.Quit(); // Quit the Excel application.
                Marshal.ReleaseComObject(worksheet); // Release the COM object for the worksheet.
                Marshal.ReleaseComObject(workbook); // Release the COM object for the workbook.
                Marshal.ReleaseComObject(excelApp); // Release the COM object for the Excel application.
            }
        }
    }
}
