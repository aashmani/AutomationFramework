using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class OpenContact
    {

        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;

        [TestMethod]
        public void TestOpenActiveContact()
        {
            try
            {
                General.Login(_xrmUri, _username, _password, this.GetType().Name);
        

                //var perf = xrmBrowser.PerformanceCenter;

                //if (!perf.IsEnabled)
                //    perf.IsEnabled = true;

                Contact.Navigate();
                Contact.OpenFirst();

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