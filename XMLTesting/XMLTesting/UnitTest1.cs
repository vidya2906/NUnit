using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;
using System.Threading;

namespace excel
{
    [TestClass]
    public class UnitTest1
    {
        IWebDriver Driver;
        IJavaScriptExecutor js;
        
        public TestContext TestContext { get; set; }

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.XML","|DataDirectory|\\XMLData.xml","user",DataAccessMethod.Sequential)]
        
        [TestMethod]
        public void ExcelData()
        {
            {
                Driver = new ChromeDriver();
                js = (IJavaScriptExecutor)Driver;
                Driver.Manage().Window.Maximize();
                Driver.Navigate().GoToUrl("https://demoqa.com/text-box");
                Driver.FindElement(By.Id("userName")).SendKeys(TestContext.DataRow[0].ToString());
                Driver.FindElement(By.Id("userEmail")).SendKeys(TestContext.DataRow[1].ToString());
                Driver.FindElement(By.Id("currentAddress")).SendKeys(TestContext.DataRow[2].ToString());
                Driver.FindElement(By.Id("permanentAddress")).SendKeys(TestContext.DataRow[3].ToString());

                js.ExecuteScript("window.scrollBy(0,250)");

                Thread.Sleep(2000);
                Driver.FindElement(By.Id("submit")).Click();
                Driver.Close();
                Driver.Quit();

            }
        }
    }
}