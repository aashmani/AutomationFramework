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
        public static XrmBrowser xrmBrowser;

        static Case()
        {
            General.GetDataFromYaml();
        }
        public static void OpenRelatedCase()
        {
            xrmBrowser.ThinkTime(500);
            xrmBrowser.Navigation.OpenRelated("Cases");
            Logs.LogHTML("Opened Relate Cases Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }

        private static void ClickNew()
        {
            xrmBrowser.Related.ClickCommand("ADD NEW CASE");
            xrmBrowser.ThinkTime(2000);
        }

        public static string Create()
        {
            ClickNew();

            var dicCreateCase = General.jsonObj.SelectToken("CreateCase");
            xrmBrowser.ThinkTime(6000);

            string Name = dicCreateCase["title"].ToString();
            string caseName = ((Name == null || Name == string.Empty) ? Name : "Test API Case");
            caseName = caseName + "_" + rnd.Next(100000, 999999).ToString();
            xrmBrowser.QuickCreate.SetValue("title", caseName);

            ClickQuickSave();

            return caseName;
        }

        private static void ClickQuickSave()
        {
            xrmBrowser.QuickCreate.Save();
            xrmBrowser.ThinkTime(10000);
            CloseDuplicateWindow();
        }

        private static void CloseDuplicateWindow()
        {

            if (xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
            {
                xrmBrowser.Dialogs.DuplicateDetection(true);
                xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Duplicate Case Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }
        }

        public static bool Search(string caseName)
        {
            xrmBrowser.Grid.Search(caseName);
            xrmBrowser.ThinkTime(1000);

            var results = xrmBrowser.Grid.GetGridItems();

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
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Case to Delete", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        public static void Delete()
        {
            try
            {
                SelectFirstCase();
                xrmBrowser.CommandBar.ClickCommand("Delete");
                xrmBrowser.ThinkTime(2000);
                xrmBrowser.Dialogs.Delete();
                Logs.LogHTML("Deleted Case Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
            }
            catch (Exception ex)
            {
                xrmBrowser.ThinkTime(1000);
                Logs.LogHTML("Delete Case Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }

        }
        public static void Close()
        {
            xrmBrowser.Dispose();
        }
        public static void Update()
        {
            var dicUpdateCase = General.jsonObj.SelectToken("UpdateCase");
            xrmBrowser.ThinkTime(6000);
            SelectFirstCase();

            xrmBrowser.Entity.SetValue(new OptionSet { Name = "caseorigincode", Value = dicUpdateCase["caseorigincode"].ToString() });
            ClickSave();
        }
        private static void ClickSave()
        {
            xrmBrowser.Entity.Save();
            xrmBrowser.ThinkTime(10000);

        }
        public static void OpenRecord()
        {
            xrmBrowser.Related.OpenGridRow(0);
            xrmBrowser.ThinkTime(2000);
        }
    }
}
