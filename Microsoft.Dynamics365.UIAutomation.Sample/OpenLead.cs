using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class OpenLead
    {

        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;
        

        [TestMethod]
        public void TestOpenActiveLead()
        {
            try
            {
                General.Login(_xrmUri, _username, _password, this.GetType().Name);

                Lead.Navigate();
                Lead.OpenRecord();
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