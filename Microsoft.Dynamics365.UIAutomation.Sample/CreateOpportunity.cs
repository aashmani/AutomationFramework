using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class CreateOpportunity
    {

        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestCreateNewOpportunity()
        {
            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();

                xrmBrowser.ThinkTime(500);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Opportunities");

                xrmBrowser.ThinkTime(200);
                xrmBrowser.Grid.SwitchView("Open Opportunities");

                xrmBrowser.ThinkTime(1000);
                xrmBrowser.CommandBar.ClickCommand("New");

                xrmBrowser.ThinkTime(5000);

                xrmBrowser.Entity.SetValue("name", "Test API Opportunity");
                xrmBrowser.Entity.SetValue("description", "Testing the create api for Opportunity");

                xrmBrowser.CommandBar.ClickCommand("Save");
                xrmBrowser.ThinkTime(2000);
            }
        }
    }
}