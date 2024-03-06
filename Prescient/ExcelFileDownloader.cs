namespace Prescient
{
    class ExcelFileDownloader
    {
        private readonly string _connectionString; // Field to store the connection string.
        private readonly HttpClient _client; // Field to hold an instance of HttpClient for making HTTP requests.

        public ExcelFileDownloader(string connectionString)
        {
            _connectionString = connectionString; // Constructor to initialize the connection string.
            _client = new HttpClient(); // Initialize HttpClient.
        }

        // Method to asynchronously download and process Excel files from the provided links.
        public async Task DownloadAndProcessExcelFilesAsync(List<string> excelLinks)
        {
            foreach (string link in excelLinks)
            {
                try
                {
                    string url = $"https://clientportal.jse.co.za{link}"; // Construct the full URL.
                    byte[] fileBytes = await _client.GetByteArrayAsync(url); // Download the Excel file as a byte array.

                    string fileName = Path.GetFileName(link); // Extract the file name from the link.
                    string directoryPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "Daily_MTM_Reports"); // Specify the directory path.
                    string filePath = Path.Combine(directoryPath, fileName); // Combine directory path and file name to get the full file path.

                    Directory.CreateDirectory(directoryPath); // Create the directory if it doesn't exist.

                    if (!File.Exists(filePath))
                    {
                        await File.WriteAllBytesAsync(filePath, fileBytes); // Write the byte array to the file.
                        Console.WriteLine($"Downloaded: {url}"); // Log the successful download.

                        await ProcessExcelFileAsync(filePath); // Process the downloaded Excel file.
                    }
                    else
                    {
                        Console.WriteLine($"File already exists: {filePath}"); // Log that the file already exists.
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error downloading or processing file: {e.Message}"); // Log any errors that occur during download or processing.
                }
            }
        }

        // Method to asynchronously process an Excel file.
        private async Task ProcessExcelFileAsync(string filePath)
        {
            try
            {
                ExcelProcessor excelProcessor = new ExcelProcessor(_connectionString); // Create an instance of ExcelProcessor.
                excelProcessor.ProcessExcelFile(filePath); // Process the Excel file.
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error processing Excel file: {e.Message}"); // Log any errors that occur during processing.
            }
        }
    }
}
