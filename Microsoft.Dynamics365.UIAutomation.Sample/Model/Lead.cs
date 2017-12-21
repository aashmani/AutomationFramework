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
        public static XrmBrowser xrmBrowser;

        static Lead()
        {
            BaseModel.GetDataFromYaml();
        }

        public static void Navigate()
        {
            xrmBrowser.Navigation.OpenSubArea("Sales", "Leads");
            Logs.LogHTML("Navigated to Leads  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            //xrmBrowser.Grid.SwitchView("All Leads");
            //Logs.LogHTML("Navigated to All Leads  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }

        private static void ClickNew()
        {
            xrmBrowser.CommandBar.ClickCommand("New");
            xrmBrowser.ThinkTime(2000);
        }
        public static string Create()
        {
            ClickNew();

            var dicCreateLead = BaseModel.jsonObj.SelectToken("CreateLead");
            xrmBrowser.ThinkTime(6000);

            string firstName = dicCreateLead["firstname"].ToString() + rnd.Next(100000, 999999).ToString();
            string lastName = dicCreateLead["lastname"].ToString();
            string displayName = firstName + " " + lastName;
            var fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = firstName },
                    new Field() {Id = "lastname", Value = lastName}
                };

            xrmBrowser.Entity.SetValue("subject", dicCreateLead["subject"].ToString());
            xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
            xrmBrowser.Entity.SetValue("mobilephone", dicCreateLead["mobilephone"].ToString());
            xrmBrowser.Entity.SetValue("description", dicCreateLead["description"].ToString());
            xrmBrowser.Entity.SetValue("emailaddress1", dicCreateLead["emailaddress1"].ToString());
            xrmBrowser.Entity.SetValue("companyname", dicCreateLead["companyname"].ToString());

            ClickSave();
            return displayName;
        }

        private static void ClickSave()
        {
            xrmBrowser.CommandBar.ClickCommand("Save & Close");
            xrmBrowser.ThinkTime(2000);
            CloseDuplicateWindow();
        }
        private static void CloseDuplicateWindow()
        {
            if (xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
            {
                xrmBrowser.Dialogs.DuplicateDetection(true);
                xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Duplicate Leads Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }
        }
        public static bool Search(string displayName)
        {
            xrmBrowser.Grid.Search(displayName);
            xrmBrowser.ThinkTime(1000);

            var results = xrmBrowser.Grid.GetGridItems();

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
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Lead", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        public static void Delete()
        {
            SelectFirstRecord();
            xrmBrowser.CommandBar.ClickCommand("Delete");
            xrmBrowser.ThinkTime(2000);
            xrmBrowser.Dialogs.Delete();
            Logs.LogHTML("Deleted Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

        }
        public static void Close()
        {
            xrmBrowser.Dispose();
        }
        public static void Update()
        {
            OpenRecord();

            var dicUpdateLead = BaseModel.jsonObj.SelectToken("UpdateLead");

            xrmBrowser.Entity.SetValue("subject", dicUpdateLead["subject"].ToString());
            xrmBrowser.Entity.SetValue("description", dicUpdateLead["description"].ToString());

            ClickSave();
        }
        public static void OpenRecord()
        {
            try
            {
                xrmBrowser.Grid.OpenRecord(0);
                xrmBrowser.ThinkTime(2000);
            }
            catch(Exception ex)
            {
                Logs.LogHTML("Open Lead Failed : " + ex.Message.Trim(), Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }

        }

    }
}
