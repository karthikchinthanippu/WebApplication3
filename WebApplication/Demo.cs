using System;

using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;
using excel = Microsoft.Office.Interop.Excel;
//using SeleniumExtras.WaitHelpers;
using System.Drawing.Imaging;

using OpenQA.Selenium.Firefox;
using System.Drawing;

namespace WebApplication
{
    class Demo
    {
        IWebDriver driver;

        [SetUp]
        public void startBrowser()
        {
            driver = new ChromeDriver("C:\\temp");

        }
        [Test]
        public void test()
        {
             driver.Navigate().GoToUrl("http://www.google.com");
            Thread.Sleep(1000);
            /* String Title = driver.Title;

             Xpath=//*[@id='hello'];
             //inpu[@id='hello1']   
             //input[contains(@id,'sub')]
             //input[contains(text(),'dude']
             //input[contains(@href,'guru99.com')]
             //input[@class='jan' and text()='kar']
             //input[starts-with(@id,'gang')]
             //input[@type='kas']//following::input 
             IWebElement element = driver.FindElement(By.XPath(""));
             element.Click();
             element.Clear();
             element.SendKeys("xis");
             Boolean x = element.Displayed;
             Boolean y = element.Enabled;
             Boolean z = element.Selected;
             element.Submit();
             String s = element.Text;
             String h = element.TagName;
            // String v = element.GetCssValue;
             SelectElement select = new SelectElement(element);
             select.SelectByText("hello");
             IList<IWebElement> options = select.Options;
             int l = options.Count;
             for (int i = 0; i < l; i++)
             {
                 String value = options.ElementAt(1).Text;
                 Console.WriteLine(value);
             }
             Boolean f = select.IsMultiple;
             select.DeselectAll();
             driver.Manage().Window.Maximize();
             driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
             

             WebDriverWait wait = new WebDriverWait(driver,TimeSpan.FromSeconds(10));
             IWebElement element2 = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("h")));

             DefaultWait<IWebDriver> fluentWait = new DefaultWait<IWebDriver>(driver);
             fluentWait.Timeout = TimeSpan.FromSeconds(10);
             fluentWait.PollingInterval = TimeSpan.FromMilliseconds(250);
             fluentWait.IgnoreExceptionTypes(typeof(NoSuchElementException));
             fluentWait.Message = "Element to be searched not found";


             var p = driver.WindowHandles.Count;  
             for(int i = 0;i<p;i++) 
             {
                 driver.SwitchTo().Window(driver.WindowHandles[p]);
             }

             driver.SwitchTo().Frame("frame name");
             driver.SwitchTo().DefaultContent();
             driver.SwitchTo().ParentFrame();*/

            //  Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            // ss.SaveAsFile("", ImageFormat.Jpeg);

          /*  excel.Application x1App = new excel.Application();
            excel.Workbook x1Workbook = x1App.Workbooks.Open(@"C:\Users\KarthikChinthanippu\Downloads\Resume\Demo.xlsx");
            excel.Worksheet x1Sheet = x1Workbook.Sheets[1];
            excel.Range x1Range = x1Sheet.UsedRange;
            String Website;
            for(int i=1;i<=2;i++)
            {
                Website = x1Range.Cells[i][1].value2;
                driver.Navigate().GoToUrl(Website);
                Thread.Sleep(30);
            }*/
        }

        [TearDown]
        public void closeBrowser()
        {
            driver.Close();
        }

    }
}
