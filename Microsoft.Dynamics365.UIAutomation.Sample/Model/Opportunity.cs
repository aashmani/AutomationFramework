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
       
        public static void Navigate()
        {

            General.xrmBrowser.ThinkTime(500);
            General.xrmBrowser.Navigation.OpenSubArea("Sales", "Opportunities");
            Logs.LogHTML("Navigated to Opportunities  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            General.xrmBrowser.ThinkTime(2000);
            General.xrmBrowser.Grid.SwitchView("Open Opportunities");
            Logs.LogHTML("Navigated to Open Opportunities  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        private static void ClickNew()
        {
            General.xrmBrowser.ThinkTime(1000);
            General.xrmBrowser.CommandBar.ClickCommand("New");
        }
        public static string Create()
        {

            ClickNew();
            var dicCreateOpportunity = General.jsonObj.SelectToken("CreateOpportunity");
            General.xrmBrowser.ThinkTime(5000);
            string oppName = dicCreateOpportunity["name"].ToString();
            oppName = ((oppName == null || oppName == string.Empty) ? oppName : "TEST_Smoke_PET_Opportunity");
            oppName = oppName + rnd.Next(100000, 999999).ToString();
            General.xrmBrowser.Entity.SetValue("name", oppName);
            General.xrmBrowser.Entity.SetValue("description", dicCreateOpportunity["description"].ToString() );

           // General.xrmBrowser.Entity.SetValue("new_deligatedto", dicCreateOpportunity["new_deligatedto"].ToString());
            ClickSave();

            return oppName;
        }
        private static void ClickSave()
        {
            General.xrmBrowser.CommandBar.ClickCommand("Save & Close");
            General.xrmBrowser.ThinkTime(5000);
            CloseDuplicateWindow();
        }
        private static void CloseDuplicateWindow()
        {

            if (General.xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
            {
                General.xrmBrowser.Dialogs.DuplicateDetection(true);
                General.xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Duplicate Opportunities Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }
        }
        public static bool Search(string Name)
        {
            General.xrmBrowser.Grid.Search(Name);
            General.xrmBrowser.ThinkTime(1000);

            var results = General.xrmBrowser.Grid.GetGridItems();

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
            General.xrmBrowser.ThinkTime(1000);
            General.xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Opportunity", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        public static void Delete()
        {
            try
            {
                SelectFirst();
                General.xrmBrowser.CommandBar.ClickCommand("Delete");
                General.xrmBrowser.ThinkTime(2000);
                General.xrmBrowser.Dialogs.Delete();
                Logs.LogHTML("Deleted Opportunity Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
            }
            catch (Exception ex)
            {
                General.xrmBrowser.ThinkTime(1000);
                Logs.LogHTML("Delete Opportunity Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }
        }
        public static void OpenFirst()
        {
            General.xrmBrowser.ThinkTime(1000);
            General.xrmBrowser.Grid.OpenRecord(0);
        }
        public static void Update()
        {
            OpenFirst();
            var dicUpdateOpportunity = General.jsonObj.SelectToken("UpdateOpportunity");
            General.xrmBrowser.Entity.SetValue("emailaddress1", dicUpdateOpportunity["emailaddress1"].ToString());
            General.xrmBrowser.Entity.SetValue("mobilephone", dicUpdateOpportunity["mobilephone"].ToString());
            General.xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse(dicUpdateOpportunity["birthdate"].ToString()));
            General.xrmBrowser.Entity.Save();

        }
      
    }
}
