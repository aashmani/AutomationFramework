using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Utility;
using OpenQA.Selenium;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class CreateOpportunity
    {

        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestCreateNewOpportunity()
        {
            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                Random rnd = new Random();
                string testCaseFile = this.GetType().Name + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss").ToString();
                Logs.LogHTML(testCaseFile, string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML(testCaseFile, "Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(500);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Opportunities");
                Logs.LogHTML(testCaseFile, "Navigated to Accounts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                
                xrmBrowser.ThinkTime(200);
                xrmBrowser.Grid.SwitchView("Open Opportunities");

                xrmBrowser.ThinkTime(1000);
                xrmBrowser.CommandBar.ClickCommand("New");

                xrmBrowser.ThinkTime(5000);

                string oppName= "Test API Opportunity" + rnd.Next(100000, 999999).ToString();
                xrmBrowser.Entity.SetValue("name", oppName);
                xrmBrowser.Entity.SetValue("description", "Testing the create api for Opportunity");

                xrmBrowser.CommandBar.ClickCommand("Save & Close");
                xrmBrowser.ThinkTime(5000);

                if (xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
                {
                    xrmBrowser.Dialogs.DuplicateDetection(true);
                    xrmBrowser.ThinkTime(2000);
                    Logs.LogHTML(testCaseFile, "Duplicate Opportunities Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                }

                xrmBrowser.Grid.Search(oppName);
                xrmBrowser.ThinkTime(1000);

                var results = xrmBrowser.Grid.GetGridItems();

                if (results.Value == null || results.Value.Count == 0)
                {
                    Logs.LogHTML(testCaseFile, "Opportunity  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }
                else
                {
                    Logs.LogHTML(testCaseFile, "Created Opportunity  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                }


                try
                {
                    xrmBrowser.ThinkTime(1000);
                    xrmBrowser.Grid.SelectRecord(0);
                    Logs.LogHTML(testCaseFile, "Selected Opportunity to Delete", Logs.HTMLSection.Details, Logs.TestStatus.Pass);


                    xrmBrowser.CommandBar.ClickCommand("Delete");
                    xrmBrowser.ThinkTime(2000);
                    xrmBrowser.Dialogs.Delete();
                    Logs.LogHTML(testCaseFile, "Deleted Opportunity Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                }
                catch (Exception ex)
                {
                    xrmBrowser.ThinkTime(1000);
                    Logs.LogHTML(testCaseFile, "Delete Opportunity ( " + oppName + " ) Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }
            }
        }
    }
}