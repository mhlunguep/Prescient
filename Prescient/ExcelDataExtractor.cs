using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace Prescient
{
    public class ExcelDataExtractor
    {
        public DataTable ExtractData(Excel.Worksheet worksheet)
        {
            DataTable dataTable = new DataTable();

            // Add columns to the DataTable
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

            try
            {
                Excel.Range range = worksheet.UsedRange;

                // Extract the full path of the workbook and the file name
                string fullPath = worksheet.Parent.FullName;
                string fileName = System.IO.Path.GetFileName(fullPath);

                // Extract the date part from the file name
                string dateString = fileName.Substring(0, 8);
                DateTime fileDate;

                // Parse the date string into a DateTime object
                if (DateTime.TryParseExact(dateString, "yyyyMMdd", null, DateTimeStyles.None, out fileDate))
                {

                }
                else
                {
                    Console.WriteLine("Failed to parse the date from the file name.");
                }

                // Start reading data from the 6th row (assuming header rows are present)
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
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting data from Excel worksheet: {ex.Message}");
            }

            return dataTable;
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
                return 0.00f;
            }
        }
    }
}
