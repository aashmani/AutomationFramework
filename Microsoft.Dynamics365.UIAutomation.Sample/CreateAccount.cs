using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Utility;
using System.Collections.Generic;
using OpenQA.Selenium;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class CreateAccount
    {
        Random rnd = new Random();
        
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestCreateNewAccount()
        {

            Guid id = Guid.NewGuid();
            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                string testCaseFile = this.GetType().Name + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss").ToString();
                Logs.LogHTML(testCaseFile, string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML(testCaseFile, "Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(500);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Accounts");
                Logs.LogHTML(testCaseFile, "Navigated to Accounts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(2000);
                xrmBrowser.Grid.SwitchView("Active Accounts");

                xrmBrowser.ThinkTime(1000);
                xrmBrowser.CommandBar.ClickCommand("New");

                xrmBrowser.ThinkTime(6000);
                string accName = "Test API Account_" + rnd.Next(100000, 999999).ToString();
                xrmBrowser.Entity.SetValue("name", accName);
                xrmBrowser.Entity.SetValue("telephone1", "555-555-3111");
                xrmBrowser.Entity.SetValue("fax", "12345678");
                xrmBrowser.Entity.SetValue("websiteurl", "https://Test.crm.dynamics.com");
                xrmBrowser.Entity.SelectLookup("parentaccountid", 2);
                xrmBrowser.Entity.SetValue("tickersymbol", "TickerSymbol3 data");
                xrmBrowser.Entity.SetValue(new OptionSet { Name = "new_typeofcustomer", Value = "Customer" });
                xrmBrowser.Entity.SetValue(new OptionSet { Name = "new_customer", Value = "Customer1" });
                xrmBrowser.Entity.SetValue("revenue", "39");
                xrmBrowser.Entity.SetValue("new_testlock", "test lock");
                xrmBrowser.Entity.SetValue("creditlimit", "10000");
                //xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse("11/1/1980"));
                var fields = new List<Field>
               {
                   new Field() { Id = "address1_line1", Value = "Test" },
                   new Field() { Id = "address1_city", Value = "Contact" },
                   new Field() { Id = "address1_postalcode", Value = "577501" }
               };
                xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "address1_composite", Fields = fields });
              
                xrmBrowser.CommandBar.ClickCommand("Save & Close");
                xrmBrowser.ThinkTime(5000);

                if (xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
                {
                    xrmBrowser.Dialogs.DuplicateDetection(true);
                    xrmBrowser.ThinkTime(2000);
                    Logs.LogHTML(testCaseFile, "Duplicate Accounts Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                }

                xrmBrowser.Grid.Search(accName);
                xrmBrowser.ThinkTime(1000);

                var results = xrmBrowser.Grid.GetGridItems();

                if (results.Value == null || results.Value.Count == 0)
                {
                    Logs.LogHTML(testCaseFile, "Account  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }
                else
                {
                    Logs.LogHTML(testCaseFile, "Created Account  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                }


                try
                {
                    xrmBrowser.ThinkTime(1000);
                    xrmBrowser.Grid.SelectRecord(0);
                    Logs.LogHTML(testCaseFile, "Selected Account to Delete", Logs.HTMLSection.Details, Logs.TestStatus.Pass);


                    xrmBrowser.CommandBar.ClickCommand("Delete");
                    xrmBrowser.ThinkTime(2000);
                    xrmBrowser.Dialogs.Delete();
                    Logs.LogHTML(testCaseFile, "Deleted Account Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                }
                catch (Exception ex)
                {
                    xrmBrowser.ThinkTime(1000);
                    Logs.LogHTML(testCaseFile, "Delete Account ( " + accName + " ) Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }
            }
        }
    }
}