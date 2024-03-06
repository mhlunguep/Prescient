using Microsoft.Extensions.Configuration;
using Prescient;
class Program
{
    static async Task Main(string[] args)
    {
        IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "appsettings.json"), optional: false, reloadOnChange: true)
            .Build();

        string connectionString = config.GetConnectionString("DefaultConnection")!;

        string url = config["URL:JCE"]!;
        HtmlContentDownloader downloader = new HtmlContentDownloader();
        List<string> excelLinks = await downloader.FetchExcelLinksAsync(url);

        ExcelFileDownloader fileDownloader = new ExcelFileDownloader(connectionString);
        await fileDownloader.DownloadAndProcessExcelFilesAsync(excelLinks);
    }
}

