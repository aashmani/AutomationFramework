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
            BaseModel.GetDataFromYaml();
        }
        static Random rnd = new Random();
        public static XrmBrowser xrmBrowser;
        public static void Navigate()
        {

            xrmBrowser.ThinkTime(500);
            xrmBrowser.Navigation.OpenSubArea("Sales", "Contacts");
            Logs.LogHTML("Navigated to Contacts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            xrmBrowser.ThinkTime(2000);
            xrmBrowser.Grid.SwitchView("Active Contacts");
            Logs.LogHTML("Navigated to Active Contacts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        private static void ClickNew()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.CommandBar.ClickCommand("New");
        }
        public static string Create()
        {

            ClickNew();
            var dicCreateContact = BaseModel.jsonObj.SelectToken("CreateContact");
            xrmBrowser.ThinkTime(5000);
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

            xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
            xrmBrowser.Entity.SetValue("emailaddress1", dicCreateContact["emailaddress1"].ToString());
            xrmBrowser.Entity.SetValue("mobilephone", dicCreateContact["mobilephone"].ToString());
            xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse(dicCreateContact["birthdate"].ToString()));
            xrmBrowser.Entity.SetValue(new OptionSet { Name = "preferredcontactmethodcode", Value = dicCreateContact["preferredcontactmethodcode"].ToString() });

            ClickSave();

            return displayName;
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
                Logs.LogHTML("Duplicate Contacts Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);

            }
        }
        public static bool Search(string Name)
        {
            xrmBrowser.Grid.Search(Name);
            xrmBrowser.ThinkTime(1000);

            var results = xrmBrowser.Grid.GetGridItems();

            if (results.Value == null || results.Value.Count == 0)
            {
                Logs.LogHTML("Contact  not found or was not created.", Logs.HTMLSection.Details, Logs.TestStatus.Fail);
                return false;
            }
            else
            {
                Logs.LogHTML("Contact  Found", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
                return false;
            }
        }

        public static void SelectFirst()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.SelectRecord(0);
            Logs.LogHTML("Selected Contact", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        public static void Delete()
        {
            try
            {
                SelectFirst();
                xrmBrowser.CommandBar.ClickCommand("Delete");
                xrmBrowser.ThinkTime(2000);
                xrmBrowser.Dialogs.Delete();
                Logs.LogHTML("Deleted Contact Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
            }
            catch (Exception ex)
            {
                xrmBrowser.ThinkTime(1000);
                Logs.LogHTML("Delete Contact Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }
        }
        public static void OpenFirst()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.OpenRecord(0);
        }
        public static void Update()
        {
            OpenFirst();
            var dicUpdateContact = BaseModel.jsonObj.SelectToken("UpdateContact");
            xrmBrowser.Entity.SetValue("emailaddress1", dicUpdateContact["emailaddress1"].ToString());
            xrmBrowser.Entity.SetValue("mobilephone", dicUpdateContact["mobilephone"].ToString());
            xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse(dicUpdateContact["birthdate"].ToString()));
            xrmBrowser.Entity.Save();

        }
        public static void Close()
        {
            xrmBrowser.Dispose();
        }
    }
}
