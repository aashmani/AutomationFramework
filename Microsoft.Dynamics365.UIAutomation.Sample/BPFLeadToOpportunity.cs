﻿using Microsoft.Dynamics365.UIAutomation.Api;
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
            Random rnd = new Random();

            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                
                Logs.LogHTML(string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML("Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(1000);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Leads");
                Logs.LogHTML("Navigated to Leads  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.CommandBar.ClickCommand("New");
                xrmBrowser.BusinessProcessFlow.SelectStage(0);
                xrmBrowser.ThinkTime(5000);

                xrmBrowser.Entity.SelectLookup("header_process_parentcontactid", 0);

                xrmBrowser.Entity.SelectLookup("header_process_parentaccountid", 0);

                xrmBrowser.Entity.SetValue(new OptionSet { Name = "header_process_purchasetimeframe", Value = "Immediate" });
                xrmBrowser.Entity.SetValue("header_process_budgetamount", "1");

                xrmBrowser.Entity.SetValue(new OptionSet { Name = "header_process_purchaseprocess",  Value = "Individual" });

                xrmBrowser.Entity.SetValue("header_process_decisionmaker");
                //xrmBrowser.Entity.SetValue("header_process_decisionmaker", "Completed");
                xrmBrowser.Entity.SetValue("header_process_description", "Test header process description");

                xrmBrowser.Entity.SetValue(new OptionSet { Name = "header_leadsourcecode", Value = "Advertisement" });

                xrmBrowser.Entity.SetValue("subject", "Test API Lead");

                string firstName = "Test" + rnd.Next(100000, 999999).ToString();
                string lastName = "Lead";
                string displayName = firstName + " " + lastName;
                var fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = firstName },
                    new Field() {Id = "lastname", Value = lastName}
                };
               
                xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
                xrmBrowser.CommandBar.ClickCommand("Save");
                xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Created Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.CommandBar.ClickCommand("Qualify");
                xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Qualified Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);


                //bool isFound = xrmBrowser.Navigation.GlobalSearch(displayName);
                //xrmBrowser.ThinkTime(2000);
                //Logs.LogHTML(testCaseFile, "Global Search Success", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
              
                //if (isFound)
                //{
                //    Logs.LogHTML(testCaseFile, "Created Opportunity From Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                //}
                //else
                //{
                //    Logs.LogHTML(testCaseFile, "Opportunity  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                //}

                //try
                //{
                //    xrmBrowser.GlobalSearch.OpenRecord("Opportunity", 0, 1000);
                //    xrmBrowser.ThinkTime(1000);
                //    Logs.LogHTML(testCaseFile, "Selected Opportunity to Delete", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                //    xrmBrowser.CommandBar.ClickCommand("Delete");
                //    xrmBrowser.ThinkTime(2000);
                //    xrmBrowser.Dialogs.Delete();
                //    Logs.LogHTML(testCaseFile, "Deleted Opportunity Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                //}
                //catch (Exception ex)
                //{
                //    xrmBrowser.ThinkTime(1000);
                //    Logs.LogHTML(testCaseFile, "Delete Opportunity ( " + displayName + " ) Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                //}
              

            }

        }
    }
}
