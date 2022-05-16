using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;

namespace ExcelDataTesting
{
    [TestClass]
    public class UnitTest1
    {
        IWebDriver driver;
        IJavaScriptExecutor jsEx;

        public TestContext TestContext { get; set; }

        [TestMethod]
        public void HandelInvalidSSLCerticateError()
        {
            ChromeOptions chromeOptions = new ChromeOptions();
            chromeOptions.AcceptInsecureCertificates = true;

            driver = new ChromeDriver(chromeOptions);
            jsEx = (IJavaScriptExecutor)driver;
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://expired.badssl.com/");
            Thread.Sleep(9000);
            driver.Close();
            driver.Quit();
        }
    }
}