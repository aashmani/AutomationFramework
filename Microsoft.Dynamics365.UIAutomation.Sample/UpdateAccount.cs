using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class UpdateAccount
    {
        //private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        //private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        //private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        public static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);


        [TestMethod]
        public void TestUpdateAccount()
        {
            Logs.LogHTML(string.Empty, Logs.HTMLSection.TestCase, Logs.TestStatus.NA, this.GetType().Name);
            Account.xrmBrowser = xrmBrowser;
            try
            {
                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML("Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                Account.Navigate();
                Account.Update();
                Account.Close();
            }
            catch (Exception ex)
            {
                Logs.LogHTML("Update  Account Failed : " + ex.Message.Trim(), Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                Helper.failedScenarios.Add(this.GetType().Name);
            }
            finally {
                Account.Close(); }
        }
    }
}