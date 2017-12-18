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
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestCreateNewAccount()
        {        

            Logs.LogHTML(string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

            
            Account.NavigateToAccounts(_xrmUri, _username, _password);
            string createdAccName = Account.CreateAccount(string.Empty);
            if (Account.SearchAccount(createdAccName))
            {
                Logs.LogHTML("Created Account Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                Account.DeleteAccount();
            }
            Account.close();
        }
    }
}