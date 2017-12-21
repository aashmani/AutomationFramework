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
    public class CreateLead
    {

        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;
        public static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);

        [TestMethod]
        public void TestCreateNewLead()
        {
            try
            {
                Random rnd = new Random();
                Lead.xrmBrowser = xrmBrowser;

                BaseModel.Login(xrmBrowser, _xrmUri, _username, _password, this.GetType().Name);

                Lead.Navigate();
                string displayName=Lead.Create();
                if (Lead.Search(displayName))
                {
                    Lead.Delete();
                }
            }
            catch(Exception ex)
            {
                Logs.LogHTML("Create Lead Failed : " + ex.Message.Trim(), Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                Helper.failedScenarios.Add(this.GetType().Name);
            }
            finally
            {
                Lead.Close();
            }          
        }
    }
}