﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
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
    public class CreateLead
    {

        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestCreateNewLead()
        {
            Random rnd = new Random();

            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                string testCaseFile = this.GetType().Name + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss").ToString();
                Logs.LogHTML(testCaseFile, string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());


                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML(testCaseFile, "Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);


                xrmBrowser.Navigation.OpenSubArea("Sales", "Leads");
                Logs.LogHTML(testCaseFile, "Navigated to Leads  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.Grid.SwitchView("All Leads");
                Logs.LogHTML(testCaseFile, "Navigated to All Leads  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.CommandBar.ClickCommand("New");
                xrmBrowser.ThinkTime(2000);

                string firstName = "Test" + rnd.Next(100000, 999999).ToString();
                string lastName = "Lead";
                string displayName = firstName + " " + lastName;
                var fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = firstName },
                    new Field() {Id = "lastname", Value = lastName}
                };
            
                xrmBrowser.Entity.SetValue("subject", "Test API Lead");
                xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
                xrmBrowser.Entity.SetValue("mobilephone", "555-555-5555");
                xrmBrowser.Entity.SetValue("description", "Test lead creation with API commands");

                xrmBrowser.CommandBar.ClickCommand("Save & Close");
                xrmBrowser.ThinkTime(2000);

                if (xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
                {
                    xrmBrowser.Dialogs.DuplicateDetection(true);
                    xrmBrowser.ThinkTime(2000);
                    Logs.LogHTML(testCaseFile, "Duplicate Leads Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                }

                xrmBrowser.Grid.Search(displayName);
                xrmBrowser.ThinkTime(1000);

                var results = xrmBrowser.Grid.GetGridItems();

                if (results.Value == null || results.Value.Count == 0)
                {
                    Logs.LogHTML(testCaseFile, "Lead  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }
                else
                {
                    Logs.LogHTML(testCaseFile, "Created Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                }


                try
                {
                    xrmBrowser.ThinkTime(1000);
                    xrmBrowser.Grid.SelectRecord(0);
                    Logs.LogHTML(testCaseFile, "Selected Lead to Delete", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                    xrmBrowser.CommandBar.ClickCommand("Delete");
                    xrmBrowser.ThinkTime(2000);
                    xrmBrowser.Dialogs.Delete();
                    Logs.LogHTML(testCaseFile, "Deleted Lead Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                }
                catch (Exception ex)
                {
                    xrmBrowser.ThinkTime(1000);
                    Logs.LogHTML(testCaseFile, "Delete Lead ( " + displayName + " ) Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }

            }
        }
    }
}