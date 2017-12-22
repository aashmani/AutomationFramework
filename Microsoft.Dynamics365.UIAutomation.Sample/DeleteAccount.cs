using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class DeleteAccount
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        public static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);

        [TestMethod]
        public void TestDeleteAccount()
        {
            Account.xrmBrowser = xrmBrowser;
            try
            {
                string name = "Test API Account";
                General.Login(xrmBrowser, _xrmUri, _username, _password, this.GetType().Name);
                if (name == string.Empty)
                {
                    Account.Navigate();
                    Account.Delete();
                }
                else
                {
                    if (Account.Search(name))
                    {
                        Account.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                General.LogError(ex.Message, this.GetType().Name);
            }
            finally
            {
                Account.Close();
            }

        }
    }
}
