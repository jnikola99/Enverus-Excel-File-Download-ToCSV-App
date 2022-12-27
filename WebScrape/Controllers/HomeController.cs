using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Net;
using WebScrape.Models;
using System.Linq;
using Microsoft.Extensions.Logging;

using Syncfusion.XlsIO;
using static System.Net.Mime.MediaTypeNames;
using HtmlAgilityPack;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Text;
using System.IO;

namespace WebScrape.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }


        private static String relativeToAbsolute(String baseUrl, String relativeUrl)
        {
            Uri baseUri = new Uri(baseUrl);
            Uri uri = new Uri(baseUri, relativeUrl);
            return uri.ToString();
        }

        private static HttpWebRequest makeRequest(string url, int timeout)
        {
            HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(url);

            req.KeepAlive = true;
            req.Accept = "application/json, text/plain, */*";
            req.Timeout = req.ReadWriteTimeout = timeout;

            req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36";
            req.Headers.Add("Accept-Language", "sr-RS,sr;q=0.8,en-US;q=0.5,en;q=0.3");
            req.Headers.Add("Connection", "keep-alive");

            return req;
        }

        private static async Task downloadFileAndSaveToLocalDiskAsync(string trueUrl,string filePath)
        {
            var request2 = makeRequest(trueUrl, 50000);
            var responseStream2 = await request2.GetResponseAsync();
            var stream2 = responseStream2.GetResponseStream();
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                stream2.CopyTo(fileStream);
            }
            
        }

        private static async Task<string> getUrlOfFileAsync(string url)
        {
            var request = makeRequest(url, 50000);
            var responseStream = await request.GetResponseAsync();
            var stream = responseStream.GetResponseStream();
            StreamReader streamReader = new StreamReader(stream);
            string html = streamReader.ReadToEnd();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            var link = doc.DocumentNode.SelectSingleNode("//a[@title='Worldwide Rig Count Nov 2022.xlsx']");
            string fileUrl = link.Attributes["href"].Value;
            return fileUrl;
        }

        private static void modifyToLast2YearsAndSaveAsCSV(string filePath)
        {
            using (ExcelEngine engine = new())
            {
                IApplication excelApp = engine.Excel;
                excelApp.DefaultVersion = ExcelVersion.Excel2016;
                excelApp.RangeIndexerMode = ExcelRangeIndexerMode.Relative;
                //Open file
                FileStream fileStream = new(filePath, FileMode.Open, FileAccess.Read);

                IWorkbook workbook = excelApp.Workbooks.Open(fileStream);

                var worksheet = workbook.Worksheets[0];
                worksheet.DeleteRow(1, 6);
                worksheet.DeleteRow(29, 690);


                fileStream.Close();

                /* Stream stream = File.Create(Path.GetFullPath(@"C:\temp\new.csv"));

                 workbook.SaveAs(stream);
                 stream.Dispose();*/
                string value = @"C:\temp\new.csv";
                using (FileStream fs = new(value, FileMode.Create))
                {
                    worksheet.SaveAs(fs, ",");
                }


            }
        }

        public async Task<IActionResult> Index()
        {
            string url = "https://bakerhughesrigcount.gcs-web.com/intl-rig-count";
            string filePath = @"C:\temp\myfile.xlsx";
         
            //Free licence of Syncfusion package
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBaFt/QHRqVVhjVFpFdEBBXHxAd1p/VWJYdVt5flBPcDwsT3RfQF9iS3xSdEVnW39ed3ZSRg==;Mgo+DSMBMAY9C3t2VVhkQlFadVdJXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxRd0dhWH1edHZVRmdbUUQ=\r\n");
            
            
            string fileUrl = await getUrlOfFileAsync(url);

            string trueUrl = relativeToAbsolute(url, fileUrl);

            await downloadFileAndSaveToLocalDiskAsync(trueUrl, filePath);
           
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