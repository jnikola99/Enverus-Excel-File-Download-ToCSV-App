using WebScrape.Controllers;
using System.Net;

namespace UnitTesting
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void shouldReturnAbsoluteURL()
        {
            string baseUrl = "https://website.com";
            string relativeUrl = "/files/basicfile.xlsx";
            Assert.AreEqual("https://website.com/files/basicfile.xlsx", HomeController.relativeToAbsolute(baseUrl, relativeUrl));
            Assert.Pass();
        }

        [Test]
        public void shouldMakeHttpRequest()
        {
            Assert.IsInstanceOf(typeof(HttpWebRequest), HomeController.makeRequest("https://google.com", 10000));
            Assert.Pass();
        }

        [Test]
        public async Task shouldReturnResponseStream()
        {
            Stream s = await HomeController.getStreamFromRequest("https://google.com");
            Assert.IsInstanceOf(typeof(Stream),s);
            Assert.Pass();
        }

    }
}