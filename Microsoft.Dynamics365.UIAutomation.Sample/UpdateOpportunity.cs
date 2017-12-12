﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api;
using System.Security;

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
                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();

                //xrmBrowser.ThinkTime(500);
                //xrmBrowser.Navigation.OpenSubArea("Sales", "Opportunities");

                //xrmBrowser.ThinkTime(200);
                //xrmBrowser.Grid.SwitchView("Open Opportunities");

                

                xrmBrowser.ThinkTime(1000);
                xrmBrowser.Navigation.GlobalSearch("PES_Testing_2");
                xrmBrowser.ThinkTime(1000);
                xrmBrowser.GlobalSearch.OpenRecord("Opportunities", 0, 1000);
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
                catch(Exception ex)
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
        }
    }
}
