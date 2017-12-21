using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Utility;
using Newtonsoft.Json;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;


namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    public class Account
    {
        static Random rnd = new Random();
        public static XrmBrowser xrmBrowser;
        static Account()
        {
            General.GetDataFromYaml();
        }


        public static void Navigate()
        {
                xrmBrowser.ThinkTime(500);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Accounts");
                Logs.LogHTML("Navigated to Accounts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(2000);
                xrmBrowser.Grid.SwitchView("Active Accounts");
        }

        private static void ClickNew()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.CommandBar.ClickCommand("New");
        }
        public static string Create()
        {
            ClickNew();
            var dicCreateAccount = General.jsonObj.SelectToken("CreateAccount");
            xrmBrowser.ThinkTime(6000);
            string Name = dicCreateAccount["name"].ToString();
            string accName = ((Name == null || Name== string.Empty) ? Name : "TEST_Smoke_PET_Account");
            xrmBrowser.Entity.SetValue("name", accName + rnd.Next(100000, 999999).ToString());
            xrmBrowser.Entity.SetValue("telephone1", dicCreateAccount["telephone1"].ToString());
            xrmBrowser.Entity.SetValue("fax", dicCreateAccount["fax"].ToString());
            xrmBrowser.Entity.SetValue("websiteurl", dicCreateAccount["websiteurl"].ToString());
            xrmBrowser.Entity.SelectLookup("parentaccountid",Convert.ToInt32(dicCreateAccount["parentaccountid"].ToString()));
            xrmBrowser.Entity.SetValue("tickersymbol", dicCreateAccount["tickersymbol"].ToString());
            xrmBrowser.Entity.SetValue(new OptionSet { Name = "new_typeofcustomer", Value = dicCreateAccount["new_typeofcustomer"].ToString() });
            xrmBrowser.Entity.SetValue(new OptionSet { Name = "new_customer", Value = dicCreateAccount["new_customer"].ToString() });
            xrmBrowser.Entity.SetValue("revenue", dicCreateAccount["revenue"].ToString());
            xrmBrowser.Entity.SetValue("new_testlock", dicCreateAccount["new_testlock"].ToString());
            xrmBrowser.Entity.SetValue("creditlimit", dicCreateAccount["creditlimit"].ToString());
            //xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse("11/1/1980"));
            var fields = new List<Field>
               {
                   new Field() { Id = "address1_line1", Value = dicCreateAccount["address1_line1"].ToString() },
                   new Field() { Id = "address1_city", Value = dicCreateAccount["address1_city"].ToString() },
                   new Field() { Id = "address1_postalcode", Value = dicCreateAccount["address1_postalcode"].ToString()}
               };
            xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "address1_composite", Fields = fields });

            ClickSave();

            return accName;
        }
        
        private static void ClickSave()
        {
            xrmBrowser.CommandBar.ClickCommand("Save & Close");
            xrmBrowser.ThinkTime(5000);
            CloseDuplicateWindow();
        }

        private static void CloseDuplicateWindow()
        {

            if (xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
            {
                xrmBrowser.Dialogs.DuplicateDetection(true);
                xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Duplicate Accounts Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }
        }
        public static bool Search( string accName)
        {
            xrmBrowser.Grid.Search(accName);
            xrmBrowser.ThinkTime(1000);

            var results = xrmBrowser.Grid.GetGridItems();

            if (results.Value == null || results.Value.Count == 0)
            {
                Logs.LogHTML("Account  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                return false;
            }
            else
                return true;
        }

        private static void SelectFirstAccount()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Account", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }

        public static void Delete()
        {
            try
            {
                SelectFirstAccount();
                xrmBrowser.CommandBar.ClickCommand("Delete");
                xrmBrowser.ThinkTime(2000);
                xrmBrowser.Dialogs.Delete();
                Logs.LogHTML("Deleted Account Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
            }
            catch (Exception ex)
            {
                xrmBrowser.ThinkTime(1000);
                Logs.LogHTML("Delete Account Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }
        }

        public static void Update() {
            OpenFirstAccount();

            var dicUpdateAccount = General.jsonObj.SelectToken("UpdateAccount");

            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Entity.SelectLookup("parentaccountid", Convert.ToInt32(dicUpdateAccount["parentaccountid"].ToString()));

            xrmBrowser.Entity.Save();
            xrmBrowser.ThinkTime(2000);

            Logs.LogHTML("Updated Account Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }

        public static  void OpenFirstAccount()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.OpenRecord(0);
        }

        public static void Close() {
            xrmBrowser.Dispose();
        }
    }
}