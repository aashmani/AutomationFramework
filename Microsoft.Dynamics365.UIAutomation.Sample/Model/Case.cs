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
    public class Case
    {
        static Random rnd = new Random();
        

        static Case()
        {
            General.GetDataFromYaml();
        }
        public static void OpenRelatedCase()
        {
            General.xrmBrowser.ThinkTime(500);
             General.xrmBrowser.Navigation.OpenRelated("Cases");
            Logs.LogHTML("Opened Relate Cases Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }

        private static void ClickNew()
        {
             General.xrmBrowser.Related.ClickCommand("ADD NEW CASE");
             General.xrmBrowser.ThinkTime(2000);
        }

        public static string Create()
        {
            ClickNew();

            var dicCreateCase = General.jsonObj.SelectToken("CreateCase");
             General.xrmBrowser.ThinkTime(6000);

            string Name = dicCreateCase["title"].ToString();
            string caseName = ((Name == null || Name == string.Empty) ? Name : "Test API Case");
            caseName = caseName + "_" + rnd.Next(100000, 999999).ToString();
             General.xrmBrowser.QuickCreate.SetValue("title", caseName);

            ClickQuickSave();

            return caseName;
        }

        private static void ClickQuickSave()
        {
             General.xrmBrowser.QuickCreate.Save();
             General.xrmBrowser.ThinkTime(10000);
            CloseDuplicateWindow();
        }

        private static void CloseDuplicateWindow()
        {

            if ( General.xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
            {
                 General.xrmBrowser.Dialogs.DuplicateDetection(true);
                 General.xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Duplicate Case Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }
        }

        public static bool Search(string caseName)
        {
             General.xrmBrowser.Grid.Search(caseName);
             General.xrmBrowser.ThinkTime(1000);

            var results =  General.xrmBrowser.Grid.GetGridItems();

            if (results.Value == null || results.Value.Count == 0)
            {
                Logs.LogHTML("Case  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                return false;
            }
            else
            {
                Logs.LogHTML("Created Case  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                return true;
            }
        }

        private static void SelectFirstCase()
        {
             General.xrmBrowser.ThinkTime(1000);
             General.xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Case to Delete", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        public static void Delete()
        {
            try
            {
                SelectFirstCase();
                 General.xrmBrowser.CommandBar.ClickCommand("Delete");
                 General.xrmBrowser.ThinkTime(2000);
                 General.xrmBrowser.Dialogs.Delete();
                Logs.LogHTML("Deleted Case Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
            }
            catch (Exception ex)
            {
                 General.xrmBrowser.ThinkTime(1000);
                Logs.LogHTML("Delete Case Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }

        }
        public static void Update()
        {
            var dicUpdateCase = General.jsonObj.SelectToken("UpdateCase");
             General.xrmBrowser.ThinkTime(6000);
            SelectFirstCase();

             General.xrmBrowser.Entity.SetValue(new OptionSet { Name = "caseorigincode", Value = dicUpdateCase["caseorigincode"].ToString() });
            ClickSave();

            Logs.LogHTML("Updated Case Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        private static void ClickSave()
        {
             General.xrmBrowser.Entity.Save();
             General.xrmBrowser.ThinkTime(10000);

        }
        public static void OpenRecord()
        {
             General.xrmBrowser.Related.OpenGridRow(0);
             General.xrmBrowser.ThinkTime(2000);
        }
    }
}
