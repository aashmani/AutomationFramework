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
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        

        [TestMethod]
        public void TestCreateNewAccount()
        {
            
            try
            {

                General.Login(_xrmUri, _username, _password, this.GetType().Name);
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
                General.LogError(ex.Message, this.GetType().Name);
            }
            finally
            {
                General.Close();
            }

        }
    }
}