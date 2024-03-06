using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
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

            try
            {
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
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing Excel file: {ex.Message}");
            }
            finally
            {
                // Close Excel objects
                workbook.Close(false);
                excelApp.Quit();

                // Release Excel objects
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
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
            dataTable.Columns.Add("Volatility", typeof(float)); 
            dataTable.Columns.Add("Delta", typeof(float)); 
            dataTable.Columns.Add("DeltaValue", typeof(float)); 
            dataTable.Columns.Add("ContractsTraded", typeof(float)); 
            dataTable.Columns.Add("OpenInterest", typeof(float)); 

            // Now the dataTable is ready with columns to hold your Excel data.

            try
            {
                // Populate DataTable with Excel data
                for (int row = 6; row <= range.Rows.Count; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    dataRow["FileDate"] = fileDate;

                    // Check if the cell value for Contract is null or empty
                    object contractValue = (range.Cells[row, 1] as Excel.Range).Value2;
                    dataRow["Contract"] = contractValue != null ? contractValue.ToString() : "";
                    double? expiryDateValue = range.Cells[row, 3].Value2 as double?;
                    dataRow["ExpiryDate"] = expiryDateValue != null ? DateTime.FromOADate((double)expiryDateValue) : DateTime.MinValue;
                    // Check if the cell value for Classification is null or empty
                    object classificationValue = (range.Cells[row, 4] as Excel.Range).Value2;
                    dataRow["Classification"] = classificationValue != null ? classificationValue.ToString() : "";                    
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

                    // Create SQL command text for MERGE statement
                    StringBuilder sqlCommandText = new StringBuilder();
                    sqlCommandText.AppendLine("MERGE INTO DailyMTM AS target");
                    sqlCommandText.AppendLine("USING (VALUES (@FileDate, @Contract, @ExpiryDate, @Classification, @Strike, @CallPut, @MTMYield, @MarkPrice, @SpotRate, @PreviousMTM, @PreviousPrice, @PremiumOnOption, @Volatility, @Delta, @DeltaValue, @ContractsTraded, @OpenInterest)) AS source (FileDate, Contract, ExpiryDate, Classification, Strike, CallPut, MTMYield, MarkPrice, SpotRate, PreviousMTM, PreviousPrice, PremiumOnOption, Volatility, Delta, DeltaValue, ContractsTraded, OpenInterest)");
                    sqlCommandText.AppendLine("ON (target.FileDate = source.FileDate AND target.Contract = source.Contract)");
                    sqlCommandText.AppendLine("WHEN NOT MATCHED BY TARGET THEN");
                    sqlCommandText.AppendLine("INSERT (FileDate, Contract, ExpiryDate, Classification, Strike, CallPut, MTMYield, MarkPrice, SpotRate, PreviousMTM, PreviousPrice, PremiumOnOption, Volatility, Delta, DeltaValue, ContractsTraded, OpenInterest)");
                    sqlCommandText.AppendLine("VALUES (source.FileDate, source.Contract, source.ExpiryDate, source.Classification, source.Strike, source.CallPut, source.MTMYield, source.MarkPrice, source.SpotRate, source.PreviousMTM, source.PreviousPrice, source.PremiumOnOption, source.Volatility, source.Delta, source.DeltaValue, source.ContractsTraded, source.OpenInterest);");

                    using (SqlCommand command = new SqlCommand(sqlCommandText.ToString(), connection))
                    {
                        // Add parameters for each column in the DataTable
                        command.Parameters.Add("@FileDate", SqlDbType.DateTime);
                        command.Parameters.Add("@Contract", SqlDbType.NVarChar, 255);
                        command.Parameters.Add("@ExpiryDate", SqlDbType.DateTime);
                        command.Parameters.Add("@Classification", SqlDbType.NVarChar, 255);
                        command.Parameters.Add("@Strike", SqlDbType.Float);
                        command.Parameters.Add("@CallPut", SqlDbType.NVarChar, 255);
                        command.Parameters.Add("@MTMYield", SqlDbType.Float);
                        command.Parameters.Add("@MarkPrice", SqlDbType.Float);
                        command.Parameters.Add("@SpotRate", SqlDbType.Float);
                        command.Parameters.Add("@PreviousMTM", SqlDbType.Float);
                        command.Parameters.Add("@PreviousPrice", SqlDbType.Float);
                        command.Parameters.Add("@PremiumOnOption", SqlDbType.Float);
                        command.Parameters.Add("@Volatility", SqlDbType.Float);
                        command.Parameters.Add("@Delta", SqlDbType.Float);
                        command.Parameters.Add("@DeltaValue", SqlDbType.Float);
                        command.Parameters.Add("@ContractsTraded", SqlDbType.Float);
                        command.Parameters.Add("@OpenInterest", SqlDbType.Float);
                        // Execute MERGE statement for each row in the DataTable
                        foreach (DataRow row in dataTable.Rows)
                        {
                            command.Parameters["@FileDate"].Value = row["FileDate"];
                            command.Parameters["@Contract"].Value = row["Contract"];
                            command.Parameters["@ExpiryDate"].Value = row["ExpiryDate"];
                            command.Parameters["@Classification"].Value = row["Classification"];
                            command.Parameters["@Strike"].Value = row["Strike"];
                            command.Parameters["@CallPut"].Value = row["CallPut"];
                            command.Parameters["@MTMYield"].Value = row["MTMYield"];
                            command.Parameters["@MarkPrice"].Value = row["MarkPrice"];
                            command.Parameters["@SpotRate"].Value = row["SpotRate"];
                            command.Parameters["@PreviousMTM"].Value = row["PreviousMTM"];
                            command.Parameters["@PreviousPrice"].Value = row["PreviousPrice"];
                            command.Parameters["@PremiumOnOption"].Value = row["PremiumOnOption"];
                            command.Parameters["@Volatility"].Value = row["Volatility"];
                            command.Parameters["@Delta"].Value = row["Delta"];
                            command.Parameters["@DeltaValue"].Value = row["DeltaValue"];
                            command.Parameters["@ContractsTraded"].Value = row["ContractsTraded"];
                            command.Parameters["@OpenInterest"].Value = row["OpenInterest"];
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing data and bulk inserting: {ex.Message}");
            }
            finally
            {
                dataTable.Dispose();
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

