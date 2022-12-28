using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Net;
using WebScrape.Models;

using Syncfusion.XlsIO;
using HtmlAgilityPack;

namespace WebScrape.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        //Returns absolute URL from a base website URL plus relativeUrl to a file or directory and returns as a string
        public static String relativeToAbsolute(String baseUrl, String relativeUrl)
        {
            Uri baseUri = new Uri(baseUrl);
            Uri uri = new Uri(baseUri, relativeUrl);
            return uri.ToString();
        }

        //Makes a Http request for a website with some specific headers we need(otherwise won't work for bakerhughes site)
        public static HttpWebRequest makeRequest(string url, int timeout)
        {
            //Make request
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(url);

            //Set header attributes
            req.KeepAlive = true;
            req.Accept = "application/json, text/plain, */*";
            req.Timeout = req.ReadWriteTimeout = timeout;

            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36";
            req.Headers.Add("Accept-Language", "sr-RS,sr;q=0.8,en-US;q=0.5,en;q=0.3");
            req.Headers.Add("Connection", "keep-alive");

            return req;
        }

        //Returns a response stream from a request
        public static async Task<Stream> getStreamFromRequest(string trueUrl)
        {
            var request = makeRequest(trueUrl, 50000);
            var responseStream = await request.GetResponseAsync();
            var stream = responseStream.GetResponseStream();
            return stream;
        } 

        //Gets stream of file and saves it to given filePath
        public static async Task downloadFileAndSaveToLocalDiskAsync(string trueUrl,string filePath)
        {
            Directory.CreateDirectory(@"C:\temp");
            Stream stream = await getStreamFromRequest(trueUrl);
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(fileStream);
            }
            
        }

        //Gets stream of html, loads it to a HtmlAgilityPack.HtmlDocument and then searches for the file we need by title
        public static async Task<string> getUrlOfFileAsync(string url)
        {
            Stream stream = await getStreamFromRequest(url);
            StreamReader streamReader = new StreamReader(stream);
            string html = streamReader.ReadToEnd();
            //Make doc and load html
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            //Search by title
            var link = doc.DocumentNode.SelectSingleNode("//a[@title='Worldwide Rig Count Nov 2022.xlsx']");
            string fileUrl = link.Attributes["href"].Value;
            //Return relative url
            return fileUrl;
        }

        //Opens .xlsx file, modifies to last 2 years and saves as CSV in same directory
        public static void modifyToLast2YearsAndSaveAsCSV(string filePath)
        {
            using (ExcelEngine engine = new())
            {
                IApplication excelApp = engine.Excel;
                excelApp.DefaultVersion = ExcelVersion.Excel2016;
                excelApp.RangeIndexerMode = ExcelRangeIndexerMode.Relative;
                //Open file
                FileStream fileStream = new(filePath, FileMode.Open, FileAccess.Read);

                IWorkbook workbook = excelApp.Workbooks.Open(fileStream);

                //Delete rows we don't need
                var worksheet = workbook.Worksheets[0];
                worksheet.DeleteRow(1, 6);
                worksheet.DeleteRow(29, 690);


                fileStream.Close();

                string value = @"C:\temp\new.csv";

                //Save file as CSV (Comma Separated Value)
                using (FileStream fs = new(value, FileMode.Create))
                {
                    worksheet.SaveAs(fs, ",");
                }


            }
        }

        //This will run first because it's specified in Program.cs
        public async Task<IActionResult> Index()
        {
            //URL of the website we need to scrape for a file and path to save that file
            string url = "https://bakerhughesrigcount.gcs-web.com/intl-rig-count";
            string filePath = @"C:\temp\myfile.xlsx";
         
            //Free licence of Syncfusion package
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBaFt/QHRqVVhjVFpFdEBBXHxAd1p/VWJYdVt5flBPcDwsT3RfQF9iS3xSdEVnW39ed3ZSRg==;Mgo+DSMBMAY9C3t2VVhkQlFadVdJXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxRd0dhWH1edHZVRmdbUUQ=\r\n");
            
            //Get relative URL of the file we need
            string fileUrl = await getUrlOfFileAsync(url);

            //Get the whole absolute URL of the file
            string trueUrl = relativeToAbsolute(url, fileUrl);

            //Download the file and save it to a local disk
            await downloadFileAndSaveToLocalDiskAsync(trueUrl, filePath);
           
            //Take last 2 years of the file and save it to a new CSV type file
            modifyToLast2YearsAndSaveAsCSV(filePath);

            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}