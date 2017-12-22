using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class UpdateContact
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        public XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);
        Random rnd = new Random();


        [TestMethod]
        public void TestUpdateContact()
        {

            Contact.xrmBrowser = xrmBrowser;
            try
            {
                //var perf = xrmBrowser.PerformanceCenter;

                //if (!perf.IsEnabled)
                //    perf.IsEnabled = true;

                General.Login(xrmBrowser, _xrmUri, _username, _password, this.GetType().Name);
                Contact.Navigate();
                Contact.Update();

            }
            catch (Exception ex)
            {
                General.LogError(ex.Message, this.GetType().Name);
            }
            finally
            {
                Contact.Close();
            }
        }
    }
}
