using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Utility;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    public class Contact
    {
        static Contact()
        {
            General.GetDataFromYaml();
        }
        static Random rnd = new Random();
            
        public static void Navigate()
        {

            General.xrmBrowser.ThinkTime(500);
            General.xrmBrowser.Navigation.OpenSubArea("Sales", "Contacts");
            Logs.LogHTML("Navigated to Contacts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            General.xrmBrowser.ThinkTime(2000);
            General.xrmBrowser.Grid.SwitchView("Active Contacts");
            Logs.LogHTML("Navigated to Active Contacts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        private static void ClickNew()
        {
            General.xrmBrowser.ThinkTime(1000);
            General.xrmBrowser.CommandBar.ClickCommand("New");
        }
        public static string Create()
        {

            ClickNew();
            var dicCreateContact = General.jsonObj.SelectToken("CreateContact");
            General.xrmBrowser.ThinkTime(5000);
            string firstName = dicCreateContact["firstName"].ToString();
            firstName = ((firstName == null || firstName == string.Empty) ? firstName : "TEST_Smoke_PET_Contact");
            firstName = firstName + rnd.Next(100000, 999999).ToString();
            string lastName = dicCreateContact["lastName"].ToString();
            string displayName = firstName + " " + lastName;
            var fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = firstName },
                    new Field() {Id = "lastname", Value = lastName}
                };

            General.xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
            General.xrmBrowser.Entity.SetValue("emailaddress1", dicCreateContact["emailaddress1"].ToString());
            General.xrmBrowser.Entity.SetValue("mobilephone", dicCreateContact["mobilephone"].ToString());
            General.xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse(dicCreateContact["birthdate"].ToString()));
            General.xrmBrowser.Entity.SetValue(new OptionSet { Name = "preferredcontactmethodcode", Value = dicCreateContact["preferredcontactmethodcode"].ToString() });

            ClickSave();

            return displayName;
        }
        private static void ClickSave()
        {
            General.xrmBrowser.CommandBar.ClickCommand("Save & Close");
            General.xrmBrowser.ThinkTime(5000);
            CloseDuplicateWindow();
        }
        private static void CloseDuplicateWindow()
        {

            if (General.xrmBrowser.Driver.IsVisible(By.Id("InlineDialog_Background")))
            {
                General.xrmBrowser.Dialogs.DuplicateDetection(true);
                General.xrmBrowser.ThinkTime(2000);
                Logs.LogHTML("Duplicate Contacts Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }
        }
        public static bool Search(string Name)
        {
            General.xrmBrowser.Grid.Search(Name);
            General.xrmBrowser.ThinkTime(1000);

            var results = General.xrmBrowser.Grid.GetGridItems();

            if (results.Value == null || results.Value.Count == 0)
            {
                Logs.LogHTML("Contact  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                return false;
            }
            else
            {
                Logs.LogHTML("Contact  Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                return true;
            }
        }

        public static void SelectFirst()
        {
            General.xrmBrowser.ThinkTime(1000);
            General.xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Contact", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        public static void Delete()
        {
            try
            {
                SelectFirst();
                General.xrmBrowser.CommandBar.ClickCommand("Delete");
                General.xrmBrowser.ThinkTime(2000);
                General.xrmBrowser.Dialogs.Delete();
                Logs.LogHTML("Deleted Contact Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
            }
            catch (Exception ex)
            {
                General.xrmBrowser.ThinkTime(1000);
                Logs.LogHTML("Delete Contact Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }
        }
        public static void OpenFirst()
        {
            General.xrmBrowser.ThinkTime(1000);
            General.xrmBrowser.Grid.OpenRecord(0);
        }
        public static void Update()
        {
            OpenFirst();
            var dicUpdateContact = General.jsonObj.SelectToken("UpdateContact");
            General.xrmBrowser.Entity.SetValue("emailaddress1", dicUpdateContact["emailaddress1"].ToString());
            General.xrmBrowser.Entity.SetValue("mobilephone", dicUpdateContact["mobilephone"].ToString());
            General.xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse(dicUpdateContact["birthdate"].ToString()));
            General.xrmBrowser.Entity.Save();
            Logs.LogHTML("Updated Contact Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
    }
}
