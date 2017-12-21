using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Utility;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    class RegisterClient
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;
        public static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);


        [TestMethod]
        public void RegisterNewClient()
        {
           Register.xrmBrowser = xrmBrowser;
            Logs.LogHTML(string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML("Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                //xrmBrowser.ThinkTime(1000);
                Register.NavigateToClientContract();

                Register.CreateNewContract();

                xrmBrowser.Dispose();
            
        }
    }
}
