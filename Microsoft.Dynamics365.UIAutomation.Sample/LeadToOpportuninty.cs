using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
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
    class LeadToOpportuninty
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void ConvertLeadToOpportuninty()
        {
            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();

                // Creating a new lead
                xrmBrowser.Navigation.OpenSubArea("Sales", "Leads");

                xrmBrowser.Grid.SwitchView("All Leads");

                xrmBrowser.CommandBar.ClickCommand("New");

                xrmBrowser.ThinkTime(2000);
                Random r = new Random();
                int n = r.Next();

                List<Field> fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = "Test"+n.ToString()},
                    new Field() {Id = "lastname", Value = "Lead"+n.ToString()}
                };
                string topic = "Test API Lead" + n.ToString();
                xrmBrowser.Entity.SetValue("subject", topic);
                xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
                xrmBrowser.Entity.SetValue("mobilephone", "555-555-5555");
                xrmBrowser.Entity.SetValue("description", "Test lead creation with API commands");
                xrmBrowser.Entity.SetValue("companyname", "Company"+n.ToString());

                xrmBrowser.CommandBar.ClickCommand("Save");
                xrmBrowser.ThinkTime(2000);

                //converting LeadToOpportuninty
                xrmBrowser.Navigation.OpenSubArea("Sales", "Leads");

                xrmBrowser.Grid.SwitchView("All Leads");
                xrmBrowser.Grid.Search(topic);
                xrmBrowser.Grid.OpenRecord(0);
                //xrmBrowser.BusinessProcessFlow.SelectStage(0);
                xrmBrowser.CommandBar.ClickCommand("Qualify");
                xrmBrowser.ThinkTime(2000);
                //Open opportunities
                xrmBrowser.Navigation.OpenSubArea("Sales", "Opportunities");
                xrmBrowser.Grid.SwitchView("Recent Opportunities");
                xrmBrowser.Grid.Search(topic);
                xrmBrowser.Grid.OpenRecord(0);
                xrmBrowser.ThinkTime(2000);

            }

        }
    }
}
