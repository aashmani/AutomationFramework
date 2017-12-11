using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class UpdateOpportunity
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestUpdateOpportunity()
        {
            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();

                xrmBrowser.ThinkTime(500);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Opportunities");

                //xrmBrowser.ThinkTime(200);
                //xrmBrowser.Grid.SwitchView("Open Opportunities");


                xrmBrowser.ThinkTime(1000);
                xrmBrowser.Grid.OpenRecord(0);
                //xrmBrowser.Entity.SetValue("identifycompetitors", true);

                //xrmBrowser.BusinessProcessFlow.SelectStage(1);

                xrmBrowser.Entity.SetValue("customerneed", "test content");
                xrmBrowser.Entity.SetValue("proposedsolution", "test content");

                
                //xrmBrowser.Entity.SetValue("identifycompetitors", true);
                //xrmBrowser.BusinessProcessFlow.SelectStage(1);
                //xrmBrowser.BusinessProcessFlow.
                //xrmBrowser.BusinessProcessFlow.SetActive();
                // 
                //Field fld = new Field();

                //xrmBrowser.Entity.SetValue(fi)
                //xrmBrowser.BusinessProcessFlow.SelectStage(2);


                xrmBrowser.Entity.SetValue("description", "Testing the update api for Opportunity");

              //  xrmBrowser.Entity.Save();
            }
        }
    }
}
