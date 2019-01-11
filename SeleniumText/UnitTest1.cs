using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace SeleniumText
{
    [TestClass]
    public class UnitTest1
    {
        static IWebDriver driverGC;

        [AssemblyInitialize]
        public static void SetUp(TestContext context)
        {
            driverGC = new ChromeDriver();
           // driverGC.Manage().Timeouts().ImplicitWait = new TimeSpan(0,0,30);
        }

        [TestMethod]
        public void TestMethod1()
        {
            driverGC.Navigate().GoToUrl("http://185.100.238.9/PRST_MM_Test_WebSite/");
            driverGC.FindElement(By.Id("Username")).SendKeys("prst");
            driverGC.FindElement(By.Id("Password")).SendKeys("12QWaszx");
            driverGC.FindElement(By.ClassName("btn-primary")).SendKeys(Keys.Enter);
            var element = driverGC.FindElement(By.ClassName("sidebar-nav-fixed"));
            Assert.IsNotNull(element);            
        }

        
        public bool TestLogin(string baseaddress, string user, string password)
        {
            driverGC.Navigate().GoToUrl(baseaddress + "/Account/LogOn");
            driverGC.FindElement(By.Id("UserName")).SendKeys(user);
            driverGC.FindElement(By.Id("Password")).SendKeys(password);
            driverGC.FindElement(By.ClassName("ui-button")).SendKeys(Keys.Enter);
            
            WebDriverWait wait = new WebDriverWait(driverGC, TimeSpan.FromSeconds(30));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Id("homeBlock")));
            var element = driverGC.FindElement(By.Id("homeBlock"));

            return element != null;
        }

        public bool TestJobOrder(string baseaddress)
        {
            
            driverGC.Navigate().GoToUrl(baseaddress + "/JobOrder");

            WebDriverWait wait = new WebDriverWait(driverGC, TimeSpan.FromSeconds(30));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Id("SearchType")));

            var oSelect = new SelectElement(driverGC.FindElement(By.Id("SearchType")));
            oSelect.SelectByValue("1");

            driverGC.FindElement(By.Name("ButtonAction:Search")).SendKeys(Keys.Enter);
            
          //  WebDriverWait wait1 = new WebDriverWait(driverGC, TimeSpan.FromSeconds(30));
          //  wait1.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.TagName("tfoot")));

            var element = driverGC.FindElement(By.TagName("tfoot"));
            //var element1 = driverGC.FindElement(By.LinkText("Dettagli"));
            
            return element != null;
        }

        public bool TestDeliveryPlans(string baseaddress)
        {
            driverGC.Navigate().GoToUrl(baseaddress + "/DeliveryPlans");

            WebDriverWait wait = new WebDriverWait(driverGC, TimeSpan.FromSeconds(30));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Id("buttonPanel")));


            
            driverGC.FindElement(By.ClassName("ui-button")).SendKeys(Keys.Enter);
            
            WebDriverWait wait1 = new WebDriverWait(driverGC, TimeSpan.FromSeconds(30));
            wait1.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.ClassName("dynatree-container")));

            var element = driverGC.FindElement(By.ClassName("dynatree-container"));
            //var element1 = driverGC.FindElement(By.LinkText("Dettagli"));

            return element != null;
        }

        public bool TestJobOrdersWorkableGroupSearch(string baseaddress)
        {
            
            driverGC.Navigate().GoToUrl(baseaddress + "/JobOrder/JobOrdersWorkableGroupSearch");

            WebDriverWait wait = new WebDriverWait(driverGC, TimeSpan.FromSeconds(30));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Id("buttonPanel")));


            
            driverGC.FindElement(By.ClassName("ui-button")).SendKeys(Keys.Enter);
            
            WebDriverWait wait1 = new WebDriverWait(driverGC, TimeSpan.FromSeconds(30));
            wait1.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.ClassName("dynatree-container")));

            var element = driverGC.FindElement(By.ClassName("dynatree-container"));
            //var element1 = driverGC.FindElement(By.LinkText("Dettagli"));

            return element != null;
        }



        [TestMethod]
        public void TestAll()
        {
            List<string> notPassedUsers = new List<string>();
            DataTable dt = GetDataTableFromCsv(@"D:\eklos\Offline\SeleniumText\SeleniumText\BV2.csv", true);

            string baseaddress = "http://localhost:27807";

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var row = dt.Rows[i];
                try
                {
                    Console.WriteLine();
                    Console.WriteLine($"Test user: {row[1].ToString()}");

                    if (!TestLogin(baseaddress, row[1].ToString(), row[2].ToString()))
                    {
                        WriteErrorLog(notPassedUsers, row[1].ToString(), $"Login Failed");
                    }
                    else
                    {
                        Console.WriteLine($"Login Passed");

                        try
                        {
                            if (!TestJobOrder(baseaddress))
                            {
                                WriteErrorLog(notPassedUsers, row[1].ToString(), $"JobOrder Failed");
                            }
                            else
                            {
                                Console.WriteLine($"JobOrder Passed");
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteErrorLog(notPassedUsers, row[1].ToString(), $"JobOrder Failed");
                        }

                        try
                        {
                            if (!TestDeliveryPlans(baseaddress))
                            {
                                WriteErrorLog(notPassedUsers, row[1].ToString(), $"DeliveryPlans Failed");
                            }
                            else
                            {
                                Console.WriteLine($"DeliveryPlans Passed");
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteErrorLog(notPassedUsers, row[1].ToString(), $"DeliveryPlans Failed");
                        }

                        try
                        {
                            if (!TestJobOrdersWorkableGroupSearch(baseaddress))
                            {
                                WriteErrorLog(notPassedUsers, row[1].ToString(), $"JobOrdersWorkableGroupSearch Failed");
                            }
                            else
                            {
                                Console.WriteLine($"JobOrdersWorkableGroupSearch Passed");
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteErrorLog(notPassedUsers, row[1].ToString(), $"JobOrdersWorkableGroupSearch Failed");
                        }
                    }
                }
                catch (Exception ex)
                {
                    notPassedUsers.Add(row[1].ToString());
                }
            }

            Console.WriteLine("notPassedUsers:" + string.Join(",", notPassedUsers));
            Assert.IsTrue(notPassedUsers.Count == 0);
        }

        private static void WriteErrorLog(List<string> notPassedUsers, string userName , string message)
        {
            Console.WriteLine(message);
            if(!notPassedUsers.Contains(userName))
                notPassedUsers.Add(userName);
        }

        static DataTable GetDataTableFromCsv(string path, bool isFirstRowHeader)
        {
            string header = isFirstRowHeader ? "Yes" : "No";

            string pathOnly = Path.GetDirectoryName(path);
            string fileName = Path.GetFileName(path);

            string sql = @"SELECT * FROM [" + fileName + "]";

            using (OleDbConnection connection = new OleDbConnection(
                      @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
                      ";Extended Properties=\"Text;HDR=" + header + "\""))
            using (OleDbCommand command = new OleDbCommand(sql, connection))
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
            {
                DataTable dataTable = new DataTable();
                dataTable.Locale = CultureInfo.CurrentCulture;
                adapter.Fill(dataTable);
                return dataTable;
            }
        }



        [AssemblyCleanup]
        public static void TearDown()
        {
            driverGC.Quit();
        }
    }
}
