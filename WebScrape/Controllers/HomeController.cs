﻿using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Net;
using WebScrape.Models;
using System.Linq;
using Microsoft.Extensions.Logging;

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
            req.Headers.Add("Referer", "https://bakerhughesrigcount.gcs-web.com/intl-rig-count");
            req.Headers.Add("traceparent", null);
            req.Headers.Add("Connection", "keep-alive");

            return req;
        }

        private static async Task downloadFileAndSaveToLocalDiskAsync(string trueUrl,string filePath)
        {
            var request2 = makeRequest(trueUrl, 50000);
            var responseStream2 = await request2.GetResponseAsync();
            var stream2 = responseStream2.GetResponseStream();
            FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            stream2.CopyTo(fileStream);
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

        public async Task<IActionResult> Index()
        {
            string url = "https://bakerhughesrigcount.gcs-web.com/intl-rig-count";
            string filePath = @"C:\temp\myfile.xlsx";
          
            string fileUrl = await getUrlOfFileAsync(url);

            string trueUrl = relativeToAbsolute(url, fileUrl);

            await downloadFileAndSaveToLocalDiskAsync(trueUrl, filePath);
           
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