using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Utility;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    class Lead
    {
        static Random rnd = new Random();
        

        static Lead()
        {
            General.GetDataFromYaml();
        }

        public static void Navigate()
        {
             General.xrmBrowser.Navigation.OpenSubArea("Sales", "Leads");
            Logs.LogHTML("Navigated to Leads  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
             General.xrmBrowser.ThinkTime(3000);
            // General.xrmBrowser.Grid.SwitchView("All Leads");
            //Logs.LogHTML("Navigated to All Leads  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }

        private static void ClickNew()
        {
             General.xrmBrowser.CommandBar.ClickCommand("New");
             General.xrmBrowser.ThinkTime(2000);
        }
        public static string Create()
        {
            ClickNew();

            var dicCreateLead = General.jsonObj.SelectToken("CreateLead");
             General.xrmBrowser.ThinkTime(6000);

            string firstName = dicCreateLead["firstname"].ToString() + rnd.Next(100000, 999999).ToString();
            string lastName = dicCreateLead["lastname"].ToString();
            string displayName = firstName + " " + lastName;
            var fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = firstName },
                    new Field() {Id = "lastname", Value = lastName}
                };

             General.xrmBrowser.Entity.SetValue("subject", dicCreateLead["subject"].ToString());
             General.xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
             General.xrmBrowser.Entity.SetValue("mobilephone", dicCreateLead["mobilephone"].ToString());
             General.xrmBrowser.Entity.SetValue("description", dicCreateLead["description"].ToString());
             General.xrmBrowser.Entity.SetValue("emailaddress1", dicCreateLead["emailaddress1"].ToString());
             General.xrmBrowser.Entity.SetValue("companyname", dicCreateLead["companyname"].ToString());

            ClickSaveClose();
            return displayName;
        }

        public static string CreateBPF()
        {
            ClickNew();
            ClickHeader();
            var dicBPFLeadToOpportunity = General.jsonObj.SelectToken("BPFLeadToOpportunity");

             General.xrmBrowser.Entity.SelectLookup("header_process_parentcontactid", Convert.ToInt32(dicBPFLeadToOpportunity["header_process_parentcontactid"].ToString()));

             General.xrmBrowser.Entity.SelectLookup("header_process_parentaccountid", Convert.ToInt32(dicBPFLeadToOpportunity["header_process_parentaccountid"].ToString()));

             General.xrmBrowser.Entity.SetValue(new OptionSet { Name = "header_process_purchasetimeframe", Value = dicBPFLeadToOpportunity["header_process_purchasetimeframe"].ToString()});
             General.xrmBrowser.Entity.SetValue("header_process_budgetamount", dicBPFLeadToOpportunity["header_process_budgetamount"].ToString());

             General.xrmBrowser.Entity.SetValue(new OptionSet { Name = "header_process_purchaseprocess", Value = dicBPFLeadToOpportunity["header_process_purchaseprocess"].ToString() });

             General.xrmBrowser.Entity.SetValue("header_process_decisionmaker");

             General.xrmBrowser.Entity.SetValue("header_process_description", dicBPFLeadToOpportunity["header_process_description"].ToString());

             General.xrmBrowser.Entity.SetValue(new OptionSet { Name = "header_leadsourcecode", Value = dicBPFLeadToOpportunity["header_leadsourcecode"].ToString() });
            string firstName = dicBPFLeadToOpportunity["firstname"].ToString() + rnd.Next(100000, 999999).ToString();
            string lastName = dicBPFLeadToOpportunity["lastname"].ToString();
            string displayName = firstName + " " + lastName;
            var fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = firstName },
                    new Field() {Id = "lastname", Value = lastName}
                };
            string subject = dicBPFLeadToOpportunity["subject"].ToString() + "_" + rnd.Next(100000, 999999).ToString();
             General.xrmBrowser.Entity.SetValue("subject", subject);
             General.xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
            ClickSave();
            return subject;
        }

        private static void ClickHeader()
        {
             General.xrmBrowser.BusinessProcessFlow.SelectStage(0);
             General.xrmBrowser.ThinkTime(3000);
        }

        private static void ClickSave()
        {
            General.xrmBrowser.Entity.Save();
            //General.xrmBrowser.CommandBar.ClickCommand("Save");
             General.xrmBrowser.ThinkTime(2000);
            CloseDuplicateWindow();
        }
        private static void ClickSaveClose()
        {
             General.xrmBrowser.CommandBar.ClickCommand("Save & Close");
             General.xrmBrowser.ThinkTime(2000);
            CloseDuplicateWindow();
        }
        private static void CloseDuplicateWindow()
        {
            if ( General.xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
            {
                 General.xrmBrowser.Dialogs.DuplicateDetection(true);
                 General.xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Duplicate Leads Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }
        }
        public static bool Search(string displayName)
        {
             General.xrmBrowser.Grid.Search(displayName);
             General.xrmBrowser.ThinkTime(1000);

            var results =  General.xrmBrowser.Grid.GetGridItems();

            if (results.Value == null || results.Value.Count == 0)
            {
                Logs.LogHTML("Lead  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                return false;
            }
            else
            {
                Logs.LogHTML("Created Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                return true;
            }

        }
        private static void SelectFirstRecord()
        {
             General.xrmBrowser.ThinkTime(1000);
             General.xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Lead", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        public static void Delete()
        {
            SelectFirstRecord();
             General.xrmBrowser.CommandBar.ClickCommand("Delete");
             General.xrmBrowser.ThinkTime(2000);
             General.xrmBrowser.Dialogs.Delete();
            Logs.LogHTML("Deleted Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

        }
        public static void Update()
        {
            OpenRecord();

            var dicUpdateLead = General.jsonObj.SelectToken("UpdateLead");

             General.xrmBrowser.Entity.SetValue("subject", dicUpdateLead["subject"].ToString());
             General.xrmBrowser.Entity.SetValue("description", dicUpdateLead["description"].ToString());

            ClickSave();

            Logs.LogHTML("Updated Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        public static void OpenRecord()
        {
            try
            {
                 General.xrmBrowser.Grid.OpenRecord(0);
                 General.xrmBrowser.ThinkTime(2000);
            }
            catch(Exception ex)
            {
                Logs.LogHTML("Open Lead Failed : " + ex.Message.Trim(), Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }

        }
        public static void ClickQualify()
        {
             General.xrmBrowser.CommandBar.ClickCommand("Qualify");
             General.xrmBrowser.ThinkTime(5000);
            Logs.LogHTML("Qualified Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

        }

    }
}
