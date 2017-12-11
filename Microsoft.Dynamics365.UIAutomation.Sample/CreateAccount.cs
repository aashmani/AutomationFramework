using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Utility;
using System.Collections.Generic;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class CreateAccount
    {

        //private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        //private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        //private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());
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

                xrmBrowser.ThinkTime(4000);
                xrmBrowser.Entity.SetValue("name", "Vinay Testing32");
                xrmBrowser.Entity.SetValue("telephone1", "555-555-3111");
                xrmBrowser.Entity.SetValue("fax", "12345678");
                xrmBrowser.Entity.SetValue("websiteurl", "https://vinayTest.crm.dynamics.com");
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
                /*Guid id = Guid.NewGuid();
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

                xrmBrowser.ThinkTime(4000);
                xrmBrowser.Entity.SetValue("name", "Test API Account_" + rnd.Next(100000, 999999).ToString());
                xrmBrowser.Entity.SetValue("telephone1", "555-555-5555");
                xrmBrowser.Entity.SetValue("websiteurl", "https://easyrepro.crm.dynamics.com");

                xrmBrowser.CommandBar.ClickCommand("Save & Close");
                xrmBrowser.ThinkTime(2000);
                Logs.LogHTML(testCaseFile, "Created Account  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
            }*/
            }
        }
    }
}