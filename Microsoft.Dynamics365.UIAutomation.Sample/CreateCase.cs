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
        public static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);

        [TestMethod]
        public void TestCreateNewCase()
        {
            Random rnd = new Random();
            try
            {
                Case.xrmBrowser = xrmBrowser;
                Account.xrmBrowser = xrmBrowser;
                BaseModel.Login(xrmBrowser, _xrmUri, _username, _password, this.GetType().Name);


                Account.Navigate();
                Account.OpenFirst();
                Case.OpenRelatedCase();
                string caseName = Case.Create();
                if (Case.Search(caseName))
                {
                    Case.Delete();
                }
            }
            catch(Exception ex)
            {
                Logs.LogHTML("Create Case Failed : " + ex.Message.Trim(), Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                Helper.failedScenarios.Add(this.GetType().Name);
            }
            finally
            {
                Case.Close();
            }
        }
    }
}
