using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class DeleteAccount
    {
        //private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        //private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        //private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestDeleteAccount()
        {
            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                
                Logs.LogHTML(string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

                string name = "Test API Account";

                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML("Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(500);
                if (name == string.Empty)
                {
                    xrmBrowser.Navigation.OpenSubArea("Sales", "Accounts");
                    Logs.LogHTML("Navigated to Accounts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                }
                else
                {
                    xrmBrowser.Navigation.GlobalSearch(name);
                    Logs.LogHTML("Global Search Success", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                }
                xrmBrowser.ThinkTime(2000);

                try
                {
                    xrmBrowser.GlobalSearch.OpenRecord("Accounts", 0, 1000);
                    xrmBrowser.ThinkTime(1000);
                    Logs.LogHTML("Selected Account to Delete", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                    xrmBrowser.CommandBar.ClickCommand("Delete");                   
                    xrmBrowser.ThinkTime(2000);
                    xrmBrowser.Dialogs.Delete();
                    Logs.LogHTML("Deleted Account Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                }
                catch (Exception ex)
                {
                    xrmBrowser.ThinkTime(1000);
                    Logs.LogHTML("Delete Account ( " + name + " ) Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }
            }
        }
    }
}

