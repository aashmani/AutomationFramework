using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Utility;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class BPFLeadToOpportunity
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestBPFLeadToOpportunity()
        {
            Random rnd = new Random();
            try
            {
                General.Login(_xrmUri, _username, _password, this.GetType().Name);

                Lead.Navigate();

                string leadName = Lead.CreateBPF();

                Lead.ClickQualify();

                Opportunity.Navigate();

                if (Opportunity.Search(leadName))
                {
                    Logs.LogHTML("Created Lead and converted to opportunity Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                    Opportunity.Delete();
                }
            }
            catch(Exception ex)
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
