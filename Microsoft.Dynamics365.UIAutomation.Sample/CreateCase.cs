using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Utility;
using OpenQA.Selenium;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class CreateCase
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestCreateNewCase()
        {
            Random rnd = new Random();

            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                
                Logs.LogHTML(string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML("Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);


                xrmBrowser.ThinkTime(500);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Accounts");
                Logs.LogHTML("Navigated to Accounts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(3000);
                xrmBrowser.Grid.OpenRecord(0);
                xrmBrowser.Navigation.OpenRelated("Cases");
                Logs.LogHTML("Opened Relate Cases Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.Related.SwitchView("Active Cases");
                xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Navigated to Active Cases  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.Related.ClickCommand("ADD NEW CASE");
                xrmBrowser.ThinkTime(2000);

                string caseName = "Test API Case_" + rnd.Next(100000, 999999).ToString();
                xrmBrowser.QuickCreate.SetValue("title", caseName);

                xrmBrowser.QuickCreate.Save();
                xrmBrowser.ThinkTime(10000);

                //xrmBrowser.CommandBar.ClickCommand("Save & Close");
                //xrmBrowser.ThinkTime(5000);

                if (xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
                {
                    xrmBrowser.Dialogs.DuplicateDetection(true);
                    xrmBrowser.ThinkTime(2000);
                    Logs.LogHTML("Duplicate Cases Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                }

                xrmBrowser.Grid.Search(caseName);
                xrmBrowser.ThinkTime(1000);

                var results = xrmBrowser.Grid.GetGridItems();

                if (results.Value == null || results.Value.Count == 0)
                {
                    Logs.LogHTML("Case  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }
                else
                {
                    Logs.LogHTML("Created Case  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                }


                try
                {
                    xrmBrowser.ThinkTime(1000);
                    xrmBrowser.Grid.SelectRecord(0);
                    Logs.LogHTML("Selected Case to Delete", Logs.HTMLSection.Details, Logs.TestStatus.Pass);


                    xrmBrowser.CommandBar.ClickCommand("Delete");
                    xrmBrowser.ThinkTime(2000);
                    xrmBrowser.Dialogs.Delete();
                    Logs.LogHTML("Deleted Case Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                }
                catch (Exception ex)
                {
                    xrmBrowser.ThinkTime(1000);
                    Logs.LogHTML("Delete Case ( " + caseName + " ) Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }

            }
        }
    }
}
