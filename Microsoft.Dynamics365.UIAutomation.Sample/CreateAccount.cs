using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Utility;
using System.Collections.Generic;
using OpenQA.Selenium;
using System.Configuration;
using YamlDotNet.Serialization;
using System.IO;
using YamlDotNet.Serialization.NamingConventions;
using Newtonsoft.Json;
using System.Text;
using System.Collections;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class CreateAccount
    {
        Random rnd = new Random();

        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        public static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);

        [TestMethod]
        public void TestCreateNewAccount()
        {

            Account.xrmBrowser = xrmBrowser;
            Logs.LogHTML(string.Empty, Logs.HTMLSection.TestCase, Logs.TestStatus.NA, this.GetType().Name);
            try
            {
                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);   
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML("Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                Account.Navigate();

                string createdAccName = Account.Create();
                if (Account.Search(createdAccName))
                {
                    Logs.LogHTML("Created Account Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                    Account.Delete();
                }
            }
            catch (Exception ex)
            {
                Logs.LogHTML("Create Account Failed : " + ex.Message.Trim(), Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                Helper.failedScenarios.Add(this.GetType().Name);
            }
            finally {
                Account.Close();
            }

        }
    }
}