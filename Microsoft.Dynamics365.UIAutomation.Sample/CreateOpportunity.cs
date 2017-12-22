using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class CreateOpportunity
    {

        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        public static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);

        [TestMethod]
        public void TestCreateNewOpportunity()
        {
            Opportunity.xrmBrowser = xrmBrowser;
            try
            {

                General.Login(xrmBrowser, _xrmUri, _username, _password, this.GetType().Name);
                Opportunity.Navigate();

                string createdName = Opportunity.Create();
                if (Opportunity.Search(createdName))
                {
                    Logs.LogHTML("Created Opportunity Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                    Opportunity.Delete();
                }
            }
            catch (Exception ex)
            {
                General.LogError(ex.Message, this.GetType().Name);
            }
            finally
            {
                Opportunity.Close();
            }

        }
    }
}