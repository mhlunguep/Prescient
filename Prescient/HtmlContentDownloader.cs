using HtmlAgilityPack;
namespace Prescient
{
    class HtmlContentDownloader
    {
        private HttpClient _client; // Field to hold an instance of HttpClient for making HTTP requests.

        public HtmlContentDownloader()
        {
            _client = new HttpClient(); // Constructor to initialize HttpClient.
        }

        // Method to fetch Excel links asynchronously from a given URL.
        public async Task<List<string>> FetchExcelLinksAsync(string url)
        {
            string htmlContent = await _client.GetStringAsync(url); // Asynchronously fetch HTML content from the specified URL.
            HtmlDocument htmlDocument = new HtmlDocument(); // Create an instance of HtmlDocument from HtmlAgilityPack.
            htmlDocument.LoadHtml(htmlContent); // Load the HTML content into the HtmlDocument.

            var linkNodes = htmlDocument.DocumentNode.SelectNodes("//a[@href]"); // Select anchor nodes with href attributes.
            List<string> excelLinks = new List<string>(); // List to store Excel links.

            if (linkNodes != null)
            {
                // Iterate through each anchor node and extract href attribute.
                foreach (var linkNode in linkNodes)
                {
                    string href = linkNode.GetAttributeValue("href", ""); // Get the value of the href attribute.

                    // Check if the link ends with .xlsx or .xls and contains "2023" in the URL.
                    if ((href.EndsWith(".xlsx") || href.EndsWith(".xls")) && href.Contains("2023"))
                    {
                        excelLinks.Add(href); // Add the link to the list of Excel links.
                    }
                }
            }

            return excelLinks; // Return the list of Excel links.
        }
    }
}
