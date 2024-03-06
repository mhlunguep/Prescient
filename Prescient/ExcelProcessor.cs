using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;
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
            Excel.Worksheet worksheet = workbook.Sheets[1];

            // Get the used range of cells in the worksheet
            Excel.Range range = worksheet.UsedRange;

            // Extract the full path of the workbook and the file name
            string fullPath = workbook.FullName;
            string fileName = System.IO.Path.GetFileName(fullPath);

            // Extract the date part from the file name
            string dateString = fileName.Substring(0, 8);
            DateTime fileDate;

            // Parse the date string into a DateTime object
            if (DateTime.TryParseExact(dateString, "yyyyMMdd", null, DateTimeStyles.None, out fileDate))
            {
                // Process data and perform bulk insertion
                ProcessDataAndBulkInsert(range, fileDate);
            }
            else
            {
                Console.WriteLine("Failed to parse the date from the file name.");
            }

            // Close Excel objects
            workbook.Close(false);
            excelApp.Quit();
        }

        private void ProcessDataAndBulkInsert(Excel.Range range, DateTime fileDate)
        {
            // Create a DataTable to hold Excel data
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FileDate", typeof(DateTime));
            dataTable.Columns.Add("Contract", typeof(string));
            dataTable.Columns.Add("ExpiryDate", typeof(DateTime));
            dataTable.Columns.Add("Classification", typeof(string));
            dataTable.Columns.Add("Strike", typeof(float));
            dataTable.Columns.Add("CallPut", typeof(string));
            dataTable.Columns.Add("MTMYield", typeof(float));
            dataTable.Columns.Add("MarkPrice", typeof(float));
            dataTable.Columns.Add("SpotRate", typeof(float));
            dataTable.Columns.Add("PreviousMTM", typeof(float));
            dataTable.Columns.Add("PreviousPrice", typeof(float));
            dataTable.Columns.Add("PremiumOnOption", typeof(float));
            dataTable.Columns.Add("Volatility", typeof(string));
            dataTable.Columns.Add("Delta", typeof(float));
            dataTable.Columns.Add("DeltaValue", typeof(float));
            dataTable.Columns.Add("ContractsTraded", typeof(float));
            dataTable.Columns.Add("OpenInterest", typeof(float));

            // Populate DataTable with Excel data
            for (int row = 6; row <= range.Rows.Count; row++)
            {
                DataRow dataRow = dataTable.NewRow();
                dataRow["FileDate"] = fileDate;

                // Check if the cell value for Contract is null or empty
                object contractValue = (range.Cells[row, 1] as Excel.Range).Value2;
                if (contractValue != null && !string.IsNullOrWhiteSpace(contractValue.ToString()))
                {
                    dataRow["Contract"] = (string)contractValue;
                }
                else
                {
                    // Handle the case where the cell value for Contract is null or empty
                    // You might set a default value or handle it based on your application logic
                    dataRow["Contract"] = ""; // Or any other default value you choose
                }

                double? expiryDateValue = range.Cells[row, 3].Value2 as double?;
                DateTime expiryDate;

                // Check if the expiryDateValue is null or not
                if (expiryDateValue != null)
                {
                    // Convert the double value to a DateTime object
                    expiryDate = DateTime.FromOADate((double)expiryDateValue);
                    // Assign the expiry date to the ExpiryDate column in the DataRow
                    dataRow["ExpiryDate"] = expiryDate;
                }
                else
                {
                    // Handle the case where the cell value is null or empty
                    // You might set a default value or handle it based on your application logic
                    dataRow["ExpiryDate"] = DateTime.MinValue; // Or any other default value you choose
                }
                // Check if the cell value for Classification is null or empty
                object classificationValue = (range.Cells[row, 4] as Excel.Range).Value2;
                if (classificationValue != null && !string.IsNullOrWhiteSpace(classificationValue.ToString()))
                {
                    dataRow["Classification"] = (string)classificationValue;
                }
                else
                {
                    // Handle the case where the cell value for Classification is null or empty
                    // You might set a default value or handle it based on your application logic
                    dataRow["Classification"] = ""; // Or any other default value you choose
                }
                

                dataRow["Strike"] = Convert.ToSingle(range.Cells[row, 5].Value2);
                dataRow["CallPut"] = GetCallPutValue(range.Cells[row, 6].Value2); // Handle empty and null values
                dataRow["MTMYield"] = Convert.ToSingle(range.Cells[row, 7].Value2);
                dataRow["MarkPrice"] = Convert.ToSingle(range.Cells[row, 8].Value2);
                dataRow["SpotRate"] = Convert.ToSingle(range.Cells[row, 9].Value2);
                dataRow["PreviousMTM"] = Convert.ToSingle(range.Cells[row, 10].Value2);
                dataRow["PreviousPrice"] = Convert.ToSingle(range.Cells[row, 11].Value2);
                dataRow["PremiumOnOption"] = Convert.ToSingle(range.Cells[row, 12].Value2);
                dataRow["Volatility"] = GetVolatilityValue(range.Cells[row, 13].Value2); // Handle 0.00 string
                dataRow["Delta"] = Convert.ToSingle(range.Cells[row, 14].Value2);
                dataRow["DeltaValue"] = Convert.ToSingle(range.Cells[row, 15].Value2);
                dataRow["ContractsTraded"] = Convert.ToSingle(range.Cells[row, 16].Value2);
                dataRow["OpenInterest"] = Convert.ToSingle(range.Cells[row, 17].Value2);


                dataTable.Rows.Add(dataRow);
            }

            // Perform bulk insertion into SQL Server
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                {
                    bulkCopy.DestinationTableName = "DailyMTM";

                    // Map DataTable columns to SQL table columns
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                    }

                    // Write data to SQL Server
                    bulkCopy.WriteToServer(dataTable);
                }
            }
        }


        // Helper method to handle empty and null values for string data types
        private string GetCallPutValue(object value)
        {
            if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
            {
                return value.ToString().Trim();
            }
            else
            {
                return null; // Return null if the value is empty or null
            }
        }

        // Helper method to handle 0.00 string for Volatility column
        private float GetVolatilityValue(object value)
        {
            if (value != null && float.TryParse(value.ToString(), out float result))
            {
                return result;
            }
            else if (value != null && value.ToString().Trim() == "0.00")
            {
                return 0.00f;
            }
            else
            {
                return 0.00f; // Default value for non-convertible or null values
            }
        }

    }
}
