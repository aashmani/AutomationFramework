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
    public class BPFLeadToOpportunity
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        //private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        //private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        //private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void TestBPFLeadToOpportunity()
        {
            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                string testCaseFile = this.GetType().Name + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss").ToString();
                Logs.LogHTML(testCaseFile, string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML(testCaseFile, "Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(1000);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Leads");
                Logs.LogHTML(testCaseFile, "Navigated to Leads  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.CommandBar.ClickCommand("New");
                xrmBrowser.BusinessProcessFlow.SelectStage(0);
                xrmBrowser.ThinkTime(1000);

                xrmBrowser.Entity.SelectLookup("header_process_parentcontactid", 0);

                xrmBrowser.Entity.SelectLookup("header_process_parentaccountid", 0);




                //xrmBrowser.Entity.Save();               
                             
                //List<Field> fields = new List<Field>
                //{
                //    new Field() {Id = "firstname", Value = "Test"},
                //    new Field() {Id = "lastname", Value = "Lead"}
                //};
                //xrmBrowser.Entity.SetValue("subject", "Test API Lead");
                //xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
                //xrmBrowser.Entity.SetValue("mobilephone", "555-555-5555");
                //xrmBrowser.Entity.SetValue("description", "Test lead creation with API commands");

                //xrmBrowser.CommandBar.ClickCommand("Save");
                //xrmBrowser.ThinkTime(2000);

                //Logs.LogHTML(testCaseFile, "Created Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }

        }
    }
}
