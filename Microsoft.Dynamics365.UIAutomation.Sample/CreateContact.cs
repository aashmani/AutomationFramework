using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Collections.Generic;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Utility;
using OpenQA.Selenium;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class CreateContact
    {

        //private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        //private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        //private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;
        Random rnd = new Random();

        [TestMethod]
        public void TestCreateNewContact()
        {
            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                string testCaseFile = this.GetType().Name + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss").ToString();
                Logs.LogHTML(testCaseFile, string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML(testCaseFile, "Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(500);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Contacts");
                Logs.LogHTML(testCaseFile, "Navigated to Contacts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(2000);
                xrmBrowser.Grid.SwitchView("Active Contacts");
                Logs.LogHTML(testCaseFile, "Navigated to Active Contacts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(1000);
                xrmBrowser.CommandBar.ClickCommand("New");

                xrmBrowser.ThinkTime(5000);

                //string firstName = "Test" + rnd.Next(100000, 999999).ToString();
                string firstName = "Test";
                string lastName = "Contact";
                string displayName = firstName + " " + lastName;
                var fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = firstName },
                    new Field() {Id = "lastname", Value = lastName}
                };

                xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
                //xrmBrowser.Entity.SetValue("emailaddress1", "test" + rnd.Next(100000, 999999).ToString() + "@contoso.com");
                xrmBrowser.Entity.SetValue("emailaddress1", "test@contoso.com");
                xrmBrowser.Entity.SetValue("mobilephone", "555-555-5555");
                xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse("11/1/1980"));
                xrmBrowser.Entity.SetValue(new OptionSet { Name = "preferredcontactmethodcode", Value = "Email" });

                xrmBrowser.CommandBar.ClickCommand("Save & Close");
                xrmBrowser.ThinkTime(5000);

                if (xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
                {
                    xrmBrowser.Dialogs.DuplicateDetection(true);
                    xrmBrowser.ThinkTime(2000);
                    Logs.LogHTML(testCaseFile, "Duplicate Contacts Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                }

                xrmBrowser.Grid.Search(displayName);
                xrmBrowser.ThinkTime(1000);
                
                var results = xrmBrowser.Grid.GetGridItems();

                if (results.Value == null || results.Value.Count == 0)
                {
                    Logs.LogHTML(testCaseFile, "Contact  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }
                else
                {
                    Logs.LogHTML(testCaseFile, "Created Contact  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                }

            }
        }
    }
}