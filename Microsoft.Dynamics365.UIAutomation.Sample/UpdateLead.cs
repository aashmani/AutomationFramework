using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class UpdateLead
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;
        

        [TestMethod]
        public void TestUpdateLead()
        {
            try
            {
                Random rnd = new Random();
                

                General.Login(_xrmUri, _username, _password, this.GetType().Name);

                Lead.Navigate();
                Lead.Update();

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
