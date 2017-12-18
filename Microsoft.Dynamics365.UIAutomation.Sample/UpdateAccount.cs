﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class UpdateAccount
    {
        //private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        //private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        //private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestUpdateAccount()
        {
            Logs.LogHTML(string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());
            General.Login(_xrmUri, _username, _password);
            Account.NavigateToAccounts();
            Account.UpdateAccount();
            Account.close();
          
        }
    }
}
