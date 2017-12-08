using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Utility;

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
                xrmBrowser.Entity.SetValue("name", "Test API Account_" + rnd.Next(100000, 999999).ToString());
                xrmBrowser.Entity.SetValue("telephone1", "555-555-5555");
                xrmBrowser.Entity.SetValue("websiteurl", "https://easyrepro.crm.dynamics.com");

                xrmBrowser.CommandBar.ClickCommand("Save & Close");
                xrmBrowser.ThinkTime(2000);
                Logs.LogHTML(testCaseFile, "Created Account  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
            }
        }
    }
}