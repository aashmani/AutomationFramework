using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Utility;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class UpdateOpportunity
    {
        private SecureString _username = string.Empty.ToSecureString();
        private readonly SecureString _password = string.Empty.ToSecureString();
        private readonly Uri _xrmUri;
        private readonly BrowserType _browser;

        [TestMethod]
        public void TestUpdateOpportunity()
        {
            using (var xrmBrowser = new XrmBrowser(TestSettings.Options))
            {
                try
                {
                    Logs.LogHTML(string.Empty, Logs.HTMLSection.Header, Logs.TestStatus.NA, this.GetType().Name, Helper.SecureStringToString(_username), _browser.ToString());

                    xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                    xrmBrowser.GuidedHelp.CloseGuidedHelp();
                    Logs.LogHTML("Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                    xrmBrowser.ThinkTime(500);
                    xrmBrowser.Navigation.OpenSubArea("Sales", "Opportunities");
                    Logs.LogHTML("Navigated to Opportunities  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                    //xrmBrowser.ThinkTime(200);
                    //xrmBrowser.Grid.SwitchView("Open Opportunities");

                    xrmBrowser.Grid.Search("PES_Testing_3");
                    xrmBrowser.ThinkTime(1000);

                    var results = xrmBrowser.Grid.GetGridItems();

                    if (results.Value != null || results.Value.Count > 0)
                    {
                        xrmBrowser.Grid.OpenRecord(0);
                        Logs.LogHTML("Opportunity Found Successfully.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);

                    }
                    else
                    {
                        Logs.LogHTML("Opportunity  not Found.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                    }

                    //    xrmBrowser.ThinkTime(1000);
                    //xrmBrowser.Navigation.GlobalSearch("PES_Testing_3");
                    //xrmBrowser.ThinkTime(1000);

                    //xrmBrowser.Grid.OpenRecord(0);
                    //xrmBrowser.Entity.SetValue("identifycompetitors", true);
                    xrmBrowser.ThinkTime(1000);
                    xrmBrowser.BusinessProcessFlow.SelectStage(0);
                    xrmBrowser.ThinkTime(1000);

                    try
                    {
                        xrmBrowser.BusinessProcessFlow.SetActive();
                    }
                    catch (Exception ex)
                    {
                        //if (ex.Message.ToLower().Contains("click"))


                    }
                    finally
                    {
                        xrmBrowser.BusinessProcessFlow.NextStage();
                    }

                    //xrmBrowser.Entity.
                    xrmBrowser.Entity.SetValue("header_process_customerneed", "test content");
                    xrmBrowser.ThinkTime(1000);
                    try
                    {
                        xrmBrowser.Entity.SetValue("header_process_proposedsolution", "test content");
                    }
                    catch (Exception ex)
                    { }
                    xrmBrowser.ThinkTime(1000);
                    xrmBrowser.Entity.SetValue("header_process_identifycompetitors");
                    xrmBrowser.Entity.SetValue("header_process_identifycustomercontacts");

                    //xrmBrowser.BusinessProcessFlow.SelectStage(2);
                    xrmBrowser.BusinessProcessFlow.NextStage();
                    xrmBrowser.ThinkTime(1000);

                    xrmBrowser.Entity.SetValue("header_process_identifypursuitteam");
                    xrmBrowser.Entity.SetValue("header_process_developproposal");
                    xrmBrowser.Entity.SetValue("header_process_completeinternalreview");
                    xrmBrowser.Entity.SetValue("header_process_presentproposal");
                    xrmBrowser.BusinessProcessFlow.NextStage();

                    //xrmBrowser.BusinessProcessFlow.SelectStage(3);
                    xrmBrowser.ThinkTime(1000);
                    xrmBrowser.Entity.SetValue("header_process_completefinalproposal");
                    xrmBrowser.Entity.SetValue("header_process_presentfinalproposal");
                    xrmBrowser.Entity.SetValue("header_process_finaldecisiondate", DateTime.Parse("12/2/1984"));
                    xrmBrowser.ThinkTime(1000);
                    xrmBrowser.Entity.SetValue("header_process_sendthankyounote");
                    xrmBrowser.Entity.SetValue("header_process_filedebrief");

                    xrmBrowser.BusinessProcessFlow.Finish();

                    Logs.LogHTML("Updated Opportunity  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                    // xrmBrowser.Entity.SetValue("identifycompetitors", "mark complete");
                    //
                    //xrmBrowser.BusinessProcessFlow.
                    //xrmBrowser.BusinessProcessFlow.SetActive();
                    // 
                    //Field fld = new Field();

                    //xrmBrowser.Entity.SetValue(fi)
                    //xrmBrowser.BusinessProcessFlow.SelectStage(2);


                    //xrmBrowser.Entity.SetValue("description", "Testing the update api for Opportunity");

                    //  xrmBrowser.Entity.Save();
                }
                catch(Exception ex)
                {
                    Logs.LogHTML("Update Opportunity Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                }
            }
        }
    }
}
