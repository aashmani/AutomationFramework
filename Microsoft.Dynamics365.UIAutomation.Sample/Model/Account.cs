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
        
        static Account()
        {
            General.GetDataFromYaml();
        }


        public static void Navigate()
        {
             General.xrmBrowser.ThinkTime(500);
             General.xrmBrowser.Navigation.OpenSubArea("Sales", "Accounts");
            Logs.LogHTML("Navigated to Accounts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

             General.xrmBrowser.ThinkTime(2000);
             General.xrmBrowser.Grid.SwitchView("Active Accounts");
        }

        private static void ClickNew()
        {
             General.xrmBrowser.ThinkTime(1000);
             General.xrmBrowser.CommandBar.ClickCommand("New");
        }
        public static string Create()
        {
            ClickNew();
            var dicCreateAccount =General.jsonObj.SelectToken("CreateAccount");
             General.xrmBrowser.ThinkTime(6000);
            string Name = dicCreateAccount["name"].ToString();
            string accName = ((Name == null || Name == string.Empty) ? Name : "TEST_Smoke_PET_Account" );
            accName = accName  +rnd.Next(100000, 999999).ToString();
             General.xrmBrowser.Entity.SetValue("name", accName);
             General.xrmBrowser.Entity.SetValue("telephone1", dicCreateAccount["telephone1"].ToString());
             General.xrmBrowser.Entity.SetValue("fax", dicCreateAccount["fax"].ToString());
             General.xrmBrowser.Entity.SetValue("websiteurl", dicCreateAccount["websiteurl"].ToString());
             General.xrmBrowser.Entity.SelectLookup("parentaccountid", Convert.ToInt32(dicCreateAccount["parentaccountid"].ToString()));
             General.xrmBrowser.Entity.SetValue("tickersymbol", dicCreateAccount["tickersymbol"].ToString());
             General.xrmBrowser.Entity.SetValue(new OptionSet { Name = "new_typeofcustomer", Value = dicCreateAccount["new_typeofcustomer"].ToString() });
             General.xrmBrowser.Entity.SetValue(new OptionSet { Name = "new_customer", Value = dicCreateAccount["new_customer"].ToString() });
             General.xrmBrowser.Entity.SetValue("revenue", dicCreateAccount["revenue"].ToString());
             General.xrmBrowser.Entity.SetValue("new_testlock", dicCreateAccount["new_testlock"].ToString());
             General.xrmBrowser.Entity.SetValue("creditlimit", dicCreateAccount["creditlimit"].ToString());
            // General.xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse("11/1/1980"));
            var fields = new List<Field>
               {
                   new Field() { Id = "address1_line1", Value = dicCreateAccount["address1_line1"].ToString() },
                   new Field() { Id = "address1_city", Value = dicCreateAccount["address1_city"].ToString() },
                   new Field() { Id = "address1_postalcode", Value = dicCreateAccount["address1_postalcode"].ToString()}
               };
             General.xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "address1_composite", Fields = fields });

            ClickSave();

            return accName;
        }

        private static void ClickSave()
        {
             General.xrmBrowser.CommandBar.ClickCommand("Save & Close");
             General.xrmBrowser.ThinkTime(5000);
            CloseDuplicateWindow();
        }

        private static void CloseDuplicateWindow()
        {

            if ( General.xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
            {
                 General.xrmBrowser.Dialogs.DuplicateDetection(true);
                 General.xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Duplicate Accounts Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }
        }
        public static bool Search(string accName)
        {
             General.xrmBrowser.Grid.Search(accName);
             General.xrmBrowser.ThinkTime(1000);

            var results =  General.xrmBrowser.Grid.GetGridItems();

            if (results.Value == null || results.Value.Count == 0)
            {
                Logs.LogHTML("Account  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                return false;
            }
            else
            {
                Logs.LogHTML("Account  Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                return true;
            }

        }

        public static void SelectFirst()
        {
             General.xrmBrowser.ThinkTime(1000);
             General.xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Account", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }

        public static void Delete()
        {
            try
            {
                SelectFirst();
                 General.xrmBrowser.CommandBar.ClickCommand("Delete");
                 General.xrmBrowser.ThinkTime(2000);
                 General.xrmBrowser.Dialogs.Delete();
                Logs.LogHTML("Deleted Account Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
            }
            catch (Exception ex)
            {
                 General.xrmBrowser.ThinkTime(1000);
                Logs.LogHTML("Delete Account Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }
        }

        public static void Update()
        {
            OpenFirst();

            var dicUpdateAccount = General.jsonObj.SelectToken("UpdateAccount");

             General.xrmBrowser.ThinkTime(1000);
             General.xrmBrowser.Entity.SelectLookup("parentaccountid", Convert.ToInt32(dicUpdateAccount["parentaccountid"].ToString()));

             General.xrmBrowser.Entity.Save();
             General.xrmBrowser.ThinkTime(2000);

            Logs.LogHTML("Updated Account Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }

        public static void OpenFirst()
        {
             General.xrmBrowser.ThinkTime(1000);
             General.xrmBrowser.Grid.OpenRecord(0);
        }
    }
}