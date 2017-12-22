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
    public class Opportunity
    {
        static Opportunity()
        {
            General.GetDataFromYaml();
        }
        static Random rnd = new Random();
        public static XrmBrowser xrmBrowser;
        public static void Navigate()
        {

            xrmBrowser.ThinkTime(500);
            xrmBrowser.Navigation.OpenSubArea("Sales", "Opportunities");
            Logs.LogHTML("Navigated to Opportunities  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            xrmBrowser.ThinkTime(2000);
            xrmBrowser.Grid.SwitchView("Open Opportunities");
            Logs.LogHTML("Navigated to Open Opportunities  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        private static void ClickNew()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.CommandBar.ClickCommand("New");
        }
        public static string Create()
        {

            ClickNew();
            var dicCreateOpportunity = General.jsonObj.SelectToken("CreateOpportunity");
            xrmBrowser.ThinkTime(5000);
            string oppName = dicCreateOpportunity["name"].ToString();
            oppName = ((oppName == null || oppName == string.Empty) ? oppName : "TEST_Smoke_PET_Opportunity");
            oppName = oppName + rnd.Next(100000, 999999).ToString();
            xrmBrowser.Entity.SetValue("name", oppName);
            xrmBrowser.Entity.SetValue("description", dicCreateOpportunity["description"].ToString() );

           // xrmBrowser.Entity.SetValue("new_deligatedto", dicCreateOpportunity["new_deligatedto"].ToString());
            ClickSave();

            return oppName;
        }
        private static void ClickSave()
        {
            xrmBrowser.CommandBar.ClickCommand("Save & Close");
            xrmBrowser.ThinkTime(5000);
            CloseDuplicateWindow();
        }
        private static void CloseDuplicateWindow()
        {

            if (xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
            {
                xrmBrowser.Dialogs.DuplicateDetection(true);
                xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Duplicate Opportunities Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }
        }
        public static bool Search(string Name)
        {
            xrmBrowser.Grid.Search(Name);
            xrmBrowser.ThinkTime(1000);

            var results = xrmBrowser.Grid.GetGridItems();

            if (results.Value == null || results.Value.Count == 0)
            {
                Logs.LogHTML("Opportunity  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                return false;
            }
            else
            {
                Logs.LogHTML("Opportunity  Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                return true;
            }
        }

        public static void SelectFirst()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Opportunity", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        public static void Delete()
        {
            try
            {
                SelectFirst();
                xrmBrowser.CommandBar.ClickCommand("Delete");
                xrmBrowser.ThinkTime(2000);
                xrmBrowser.Dialogs.Delete();
                Logs.LogHTML("Deleted Opportunity Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
            }
            catch (Exception ex)
            {
                xrmBrowser.ThinkTime(1000);
                Logs.LogHTML("Delete Opportunity Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }
        }
        public static void OpenFirst()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.OpenRecord(0);
        }
        public static void Update()
        {
            OpenFirst();
            var dicUpdateOpportunity = General.jsonObj.SelectToken("UpdateOpportunity");
            xrmBrowser.Entity.SetValue("emailaddress1", dicUpdateOpportunity["emailaddress1"].ToString());
            xrmBrowser.Entity.SetValue("mobilephone", dicUpdateOpportunity["mobilephone"].ToString());
            xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse(dicUpdateOpportunity["birthdate"].ToString()));
            xrmBrowser.Entity.Save();

        }
        public static void Close()
        {
            xrmBrowser.Dispose();
        }
    }
}
