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
using YamlDotNet.Serialization;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    public class Account
    {
        static Random rnd = new Random();
        static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);
        static Dictionary<string, object> dictionaryMain = new Dictionary<string, object>();
        static Dictionary<string, string> dictionaryCreateAccount = new Dictionary<string, string>();
        static Dictionary<string, string> dictionaryUpdateAccount = new Dictionary<string, string>();
        static Account()
        {
            try
            {
                string filepath = ConfigurationManager.AppSettings["TestDetailsYAML"].ToString();
                var reader = new StreamReader(filepath);
                var deserializer = new DeserializerBuilder().Build();
                var yamlObject = deserializer.Deserialize(reader);

                var serializer = new SerializerBuilder()
                    .JsonCompatible()
                    .Build();

                var json = serializer.Serialize(yamlObject);
                dictionaryMain = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                dictionaryCreateAccount = JsonConvert.DeserializeObject<Dictionary<string, string>>(dictionaryMain["CreateAccount"].ToString());
                dictionaryUpdateAccount = JsonConvert.DeserializeObject<Dictionary<string, string>>(dictionaryMain["UpdateAccount"].ToString());
            }
            catch(Exception ex)
            {

            }
        }



        public static void NavigateToAccounts(Uri uri, SecureString username, SecureString password)
        {
           
                xrmBrowser.LoginPage.Login(uri, username, password);
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                Logs.LogHTML("Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(500);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Accounts");
                Logs.LogHTML("Navigated to Accounts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

                xrmBrowser.ThinkTime(2000);
                xrmBrowser.Grid.SwitchView("Active Accounts");
        }
        private static void NewAccount()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.CommandBar.ClickCommand("New");
        }
        public static string CreateAccount()
        {
            NewAccount();

            xrmBrowser.ThinkTime(6000);
            string Name = dictionaryCreateAccount["name"].ToString();
            string accName = ((Name == null || Name== string.Empty) ? Name : "TEST_Smoke_PET_Account");
            xrmBrowser.Entity.SetValue("name", accName + rnd.Next(100000, 999999).ToString());
            xrmBrowser.Entity.SetValue("telephone1", dictionaryCreateAccount["telephone1"].ToString());
            xrmBrowser.Entity.SetValue("fax", dictionaryCreateAccount["fax"].ToString());
            xrmBrowser.Entity.SetValue("websiteurl", dictionaryCreateAccount["websiteurl"].ToString());
            xrmBrowser.Entity.SelectLookup("parentaccountid",Convert.ToInt32(dictionaryCreateAccount["parentaccountid"].ToString()));
            xrmBrowser.Entity.SetValue("tickersymbol", dictionaryCreateAccount["tickersymbol"].ToString());
            xrmBrowser.Entity.SetValue(new OptionSet { Name = "new_typeofcustomer", Value = dictionaryCreateAccount["new_typeofcustomer"].ToString() });
            xrmBrowser.Entity.SetValue(new OptionSet { Name = "new_customer", Value = dictionaryCreateAccount["new_customer"].ToString() });
            xrmBrowser.Entity.SetValue("revenue", dictionaryCreateAccount["revenue"].ToString());
            xrmBrowser.Entity.SetValue("new_testlock", dictionaryCreateAccount["new_testlock"].ToString());
            xrmBrowser.Entity.SetValue("creditlimit", dictionaryCreateAccount["creditlimit"].ToString());
            //xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse("11/1/1980"));
            var fields = new List<Field>
               {
                   new Field() { Id = "address1_line1", Value = dictionaryCreateAccount["address1_line1"].ToString() },
                   new Field() { Id = "address1_city", Value = dictionaryCreateAccount["address1_city"].ToString() },
                   new Field() { Id = "address1_postalcode", Value = dictionaryCreateAccount["address1_postalcode"].ToString()}
               };
            xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "address1_composite", Fields = fields });

            SaveAccount();

            return accName;
        }
        
        private static void SaveAccount()
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
        public static bool SearchAccount( string accName)
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

        public static void SelectFirstAccount()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Account", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }

        public static void DeleteAccount()
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
        public static void UpdateAccount() {
            OpenFirstAccount();

            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Entity.SelectLookup("parentaccountid", Convert.ToInt32(dictionaryUpdateAccount["parentaccountid"].ToString()));

            xrmBrowser.Entity.Save();
            xrmBrowser.ThinkTime(2000);
        }

        public static  void OpenFirstAccount()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.OpenRecord(0);
        }

        public static void close() {
            xrmBrowser.Dispose();
        }
    }
}