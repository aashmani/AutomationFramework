using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class UpdateCase
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;
        public static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);

        [TestMethod]
        public void TestUpdateCase()
        {
            try
            {
                Case.xrmBrowser = xrmBrowser;
                Account.xrmBrowser = xrmBrowser;

                General.Login(xrmBrowser, _xrmUri, _username, _password, this.GetType().Name);
                Account.Navigate();
                Account.OpenFirst();
                Case.OpenRelatedCase();
                Case.Update();
            }
            catch(Exception ex)
            {

                Logs.LogHTML("Update Case Failed : " + ex.Message.Trim(), Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                Helper.failedScenarios.Add(this.GetType().Name);
            }
            finally
            {
                Case.Close();
            }
        }
    }
}
