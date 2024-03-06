using HtmlAgilityPack;
using Microsoft.Extensions.Configuration;
using Prescient;

class Program
{
    static async Task Main(string[] args)
    {
        IConfiguration config = new ConfigurationBuilder()
    .AddJsonFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "appsettings.json"), optional: false, reloadOnChange: true)
    .Build();


        string connectionString = config.GetConnectionString("DefaultConnection");

        string url = "https://clientportal.jse.co.za/downloadable-files?RequestNode=/YieldX/Derivatives/Docs_DMTM";

        string htmlContent = await FetchHtmlContentAsync(url);
        List<string> excelLinks = ExtractExcelLinks(htmlContent);

        await DownloadAndProcessExcelFilesAsync(excelLinks, connectionString);
    }

    static async Task<string> FetchHtmlContentAsync(string url)
    {
        using (HttpClient client = new HttpClient())
        {
            return await client.GetStringAsync(url);
        }
    }

    static List<string> ExtractExcelLinks(string htmlContent)
    {
        List<string> excelLinks = new List<string>();

        HtmlDocument htmlDocument = new HtmlDocument();
        htmlDocument.LoadHtml(htmlContent);

        var linkNodes = htmlDocument.DocumentNode.SelectNodes("//a[@href]");
        if (linkNodes != null)
        {
            foreach (var linkNode in linkNodes)
            {
                string href = linkNode.GetAttributeValue("href", "");
                if ((href.EndsWith(".xlsx") || href.EndsWith(".xls")) && href.Contains("2023"))
                {
                    excelLinks.Add(href);
                }
            }
        }

        return excelLinks;
    }

    static async Task DownloadAndProcessExcelFilesAsync(List<string> excelLinks, string connectionString)
    {
        string domain = "https://clientportal.jse.co.za";
        string directoryPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "Daily_MTM_Reports");

        Directory.CreateDirectory(directoryPath);

        using (HttpClient client = new HttpClient())
        {
            foreach (string link in excelLinks)
            {
                try
                {
                    string url = $"{domain}{link}";
                    byte[] fileBytes = await client.GetByteArrayAsync(url);

                    string fileName = Path.GetFileName(link);
                    string filePath = Path.Combine(directoryPath, fileName);

                    if (!File.Exists(filePath))
                    {
                        await File.WriteAllBytesAsync(filePath, fileBytes);
                        Console.WriteLine($"Downloaded: {url}");
                    }
                    else
                    {
                        Console.WriteLine($"File already exists: {filePath}");
                    }

                    await ProcessExcelFileAsync(filePath, connectionString);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error downloading or processing file: {e.Message}");
                }
            }
        }
    }

    static async Task ProcessExcelFileAsync(string filePath, string connectionString)
    {
        try
        {
            // Use Prescient library to process Excel file
            ExcelProcessor excelProcessor = new ExcelProcessor(connectionString);
            excelProcessor.ProcessExcelFile(filePath);
        }
        catch (Exception e)
        {
            Console.WriteLine($"Error processing Excel file: {e.Message}");
        }
    }
}
