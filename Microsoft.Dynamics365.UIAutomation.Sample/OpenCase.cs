using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class OpenCase
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;
        public static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);

        [TestMethod]
        public void TestOpenCase()
        {
            try
            {

                Case.xrmBrowser = xrmBrowser;
                Account.xrmBrowser = xrmBrowser;

                BaseModel.Login(xrmBrowser, _xrmUri, _username, _password, this.GetType().Name);

                Account.Navigate();
                Account.OpenFirst();
                Case.OpenRelatedCase();
                Case.OpenRecord();
            }
            catch(Exception ex)
            {
                Logs.LogHTML("Open Case Failed : " + ex.Message.Trim(), Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                Helper.failedScenarios.Add(this.GetType().Name);
            }
            finally
            {
                Case.Close();
            }
            //using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            //{
            //    xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);

            //    xrmBrowser.GuidedHelp.CloseGuidedHelp();

            //    xrmBrowser.ThinkTime(500);
            //    xrmBrowser.Navigation.OpenSubArea("Sales", "Accounts");

            //    xrmBrowser.ThinkTime(3000);
            //    xrmBrowser.Grid.OpenRecord(0);
            //    xrmBrowser.Navigation.OpenRelated("Cases");

            //    xrmBrowser.Related.SwitchView("Active Cases");

            //    xrmBrowser.ThinkTime(2000);
            //    xrmBrowser.Related.OpenGridRow(0);
            //    xrmBrowser.ThinkTime(2000);


            //}
        }
    }
}
