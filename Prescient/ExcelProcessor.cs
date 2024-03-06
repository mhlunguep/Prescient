using Microsoft.Data.SqlClient;
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

            // Get the full path of the workbook
            string fullPath = workbook.FullName;

            // Extract the file name from the full path
            string fileName = System.IO.Path.GetFileName(fullPath);
            // Extract the date part from the file name
            string dateString = fileName.Substring(0, 8);

            string FileDate = ""; 

            // Parse the date string into a DateTime object
            DateTime fileDate;
            if (DateTime.TryParseExact(dateString, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out fileDate))
            {
                // Format the date as "yyyy-MM-dd"
                FileDate = fileDate.ToString("yyyy-MM-dd");
            }
            else
            {
                Console.WriteLine("Failed to parse the date from the file name.");
            }
        

            // Connect to the SQL database
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                for (int row = 6; row <= range.Rows.Count; row++) 
                {
                    // Extract data from Excel
                    string contract = (string)(range.Cells[row, 1] as Excel.Range).Value2;
                    double expiryDateValue = (double)(range.Cells[row, 3] as Excel.Range).Value2;
                    DateTime expiryDate = DateTime.FromOADate(expiryDateValue);
                    string classification = (string)(range.Cells[row, 4] as Excel.Range).Value2;
                    float strike = (float)(range.Cells[row, 5] as Excel.Range).Value2;
                    string callPut = (string)(range.Cells[row, 6] as Excel.Range).Value2;
                    float mtmYield = (float)(range.Cells[row, 7] as Excel.Range).Value2;
                    float markPrice = (float)(range.Cells[row, 8] as Excel.Range).Value2;
                    float spotRate = (float)(range.Cells[row, 9] as Excel.Range).Value2;
                    float previousMTM = (float)(range.Cells[row, 10] as Excel.Range).Value2;
                    float previousPrice = (float)(range.Cells[row, 11] as Excel.Range).Value2;
                    float premiumOnOption = (float)(range.Cells[row, 12] as Excel.Range).Value2;
                    string volatility = (string)(range.Cells[row, 13] as Excel.Range).Value2;
                    float delta = (float)(range.Cells[row, 14] as Excel.Range).Value2;
                    float deltaValue = (float)(range.Cells[row, 15] as Excel.Range).Value2;
                    float contractsTraded = (float)(range.Cells[row, 16] as Excel.Range).Value2;
                    float openInterest = (float)(range.Cells[row, 17] as Excel.Range).Value2;

                    // Insert contract details into the database
                    InsertContractDetails(connection, fileDate, contract, expiryDate, classification, strike, callPut, mtmYield, markPrice, spotRate, previousMTM, previousPrice, premiumOnOption, volatility, delta, deltaValue, contractsTraded, openInterest);
                }
            }

            // Close Excel objects
            workbook.Close(false);
            excelApp.Quit();
        }

        private void InsertContractDetails(SqlConnection connection, DateTime fileDate, string contract, DateTime expiryDate, string classification, float strike, string callPut, float mtmYield, float markPrice, float spotRate, float previousMTM, float previousPrice, float premiumOnOption, string volatility, float delta, float deltaValue, float contractsTraded, float openInterest)
        {
            string query = @"IF NOT EXISTS (SELECT 1 FROM DailyMTM WHERE Contract = @Contract AND ExpiryDate = @ExpiryDate AND FileDate = @FileDate)
                    BEGIN
                        INSERT INTO DailyMTM (Contract, ExpiryDate, Classification, Strike, CallPut, MTMYield, MarkPrice, SpotRate, PreviousMTM, PreviousPrice, PremiumOnOption, Volatility, Delta, DeltaValue, ContractsTraded, OpenInterest, FileDate) 
                        VALUES (@Contract, @ExpiryDate, @Classification, @Strike, @CallPut, @MTMYield, @MarkPrice, @SpotRate, @PreviousMTM, @PreviousPrice, @PremiumOnOption, @Volatility, @Delta, @DeltaValue, @ContractsTraded, @OpenInterest, @FileDate)
                    END";

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@FileDate", fileDate);
                command.Parameters.AddWithValue("@Contract", contract);
                command.Parameters.AddWithValue("@ExpiryDate", expiryDate);
                command.Parameters.AddWithValue("@Classification", classification);
                command.Parameters.AddWithValue("@Strike", strike);

                // Check for null value of callPut
                if (callPut != null)
                {
                    command.Parameters.AddWithValue("@CallPut", callPut);
                }
                else
                {
                    // Handle null value, you can set it to DBNull.Value or a default value as per your requirement
                    command.Parameters.AddWithValue("@CallPut", DBNull.Value); // or command.Parameters.AddWithValue("@CallPut", "DefaultValue");
                }

                command.Parameters.AddWithValue("@MTMYield", mtmYield);
                command.Parameters.AddWithValue("@MarkPrice", markPrice);
                command.Parameters.AddWithValue("@SpotRate", spotRate);
                command.Parameters.AddWithValue("@PreviousMTM", previousMTM);
                command.Parameters.AddWithValue("@PreviousPrice", previousPrice);
                command.Parameters.AddWithValue("@PremiumOnOption", premiumOnOption);
                command.Parameters.AddWithValue("@Volatility", volatility);
                command.Parameters.AddWithValue("@Delta", delta);
                command.Parameters.AddWithValue("@DeltaValue", deltaValue);
                command.Parameters.AddWithValue("@ContractsTraded", contractsTraded);
                command.Parameters.AddWithValue("@OpenInterest", openInterest);

                command.ExecuteNonQuery();
            }
        }
    }
}
