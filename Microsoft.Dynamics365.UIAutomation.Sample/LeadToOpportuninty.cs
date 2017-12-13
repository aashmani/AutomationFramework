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
                Logs.LogHTML(string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

                try
                {
                    xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                    xrmBrowser.GuidedHelp.CloseGuidedHelp();
                    Logs.LogHTML("Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                    // Creating a new lead
                    xrmBrowser.Navigation.OpenSubArea("Sales", "Leads");

                    xrmBrowser.Grid.SwitchView("All Leads");
                    Logs.LogHTML("Navigated to Leads  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                    xrmBrowser.CommandBar.ClickCommand("New");

                    xrmBrowser.ThinkTime(2000);
                    Random rnd = new Random();
                    string rndString = rnd.Next(100000, 999999).ToString();

                    List<Field> fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = "Test"+rndString},
                    new Field() {Id = "lastname", Value = "Lead"+rndString}
                };
                    string topic = "Test API Lead" + rndString;
                    xrmBrowser.Entity.SetValue("subject", topic);
                    xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
                    xrmBrowser.Entity.SetValue("mobilephone", "555-555-5555");
                    xrmBrowser.Entity.SetValue("description", "Test lead creation with API commands");
                    xrmBrowser.Entity.SetValue("companyname", "Company" + rndString);

                    xrmBrowser.CommandBar.ClickCommand("Save");
                    xrmBrowser.ThinkTime(2000);


                    //converting LeadToOpportuninty
                    xrmBrowser.Navigation.OpenSubArea("Sales", "Leads");

                    xrmBrowser.Grid.SwitchView("All Leads");
                    xrmBrowser.Grid.Search(topic);

                    var results = xrmBrowser.Grid.GetGridItems();

                    if (results.Value == null || results.Value.Count == 0)
                    {
                        Logs.LogHTML("Lead  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                    }
                    else
                    {
                        Logs.LogHTML("Lead Created Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                    }
                    xrmBrowser.Grid.OpenRecord(0);
                    xrmBrowser.CommandBar.ClickCommand("Qualify");
                    xrmBrowser.ThinkTime(2000);

                    //Open opportunities
                    xrmBrowser.Navigation.OpenSubArea("Sales", "Opportunities");
                    xrmBrowser.Grid.SwitchView("Recent Opportunities");
                    xrmBrowser.Grid.Search(topic);

                    var resultsopp = xrmBrowser.Grid.GetGridItems();

                    if (resultsopp.Value == null || resultsopp.Value.Count == 0)
                    {
                        Logs.LogHTML("Opportunity  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                    }
                    else
                    {
                        Logs.LogHTML("Lead Converted to Oppotunity Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                    }

                    xrmBrowser.Grid.OpenRecord(0);
                    xrmBrowser.ThinkTime(2000);
                }
                catch (Exception ex)
                {
                    Logs.LogHTML("Lead to Opportunity Convertion Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }
            }

        }

    }
}
