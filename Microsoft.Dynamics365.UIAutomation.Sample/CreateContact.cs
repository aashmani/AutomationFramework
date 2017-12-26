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
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        Random rnd = new Random();

        [TestMethod]
        public void TestCreateNewContact()
        {
            try
            {
               
                General.Login(_xrmUri, _username, _password, this.GetType().Name);
                Contact.Navigate();

                string createdName = Contact.Create();
                if (Contact.Search(createdName))
                {
                    Logs.LogHTML("Contact Created Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                    Contact.Delete();
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