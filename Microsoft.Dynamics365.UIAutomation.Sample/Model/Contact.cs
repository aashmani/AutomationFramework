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

            xrmBrowser.ThinkTime(5000);

            string firstName = "Test" + rnd.Next(100000, 999999).ToString();
            //string firstName = "Test";
            string lastName = "Contact";
            string displayName = firstName + " " + lastName;
            var fields = new List<Field>
                {
                    new Field() {Id = "firstname", Value = firstName },
                    new Field() {Id = "lastname", Value = lastName}
                };

            xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
            xrmBrowser.Entity.SetValue("emailaddress1", "test" + rnd.Next(100000, 999999).ToString() + "@contoso.com");
            //xrmBrowser.Entity.SetValue("emailaddress1", "test@contoso.com");
            xrmBrowser.Entity.SetValue("mobilephone", "555-555-5555");
            xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse("11/1/1980"));
            xrmBrowser.Entity.SetValue(new OptionSet { Name = "preferredcontactmethodcode", Value = "Email" });

            ClickSave();

            return firstName;
        }
        private static void ClickSave()
        {
            xrmBrowser.CommandBar.ClickCommand("Save & Close");
            xrmBrowser.ThinkTime(5000);
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
            xrmBrowser.Entity.SetValue("emailaddress1", "testUpdate@contoso.com");
            xrmBrowser.Entity.SetValue("mobilephone", "123-222-4444");
            xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse("12/2/1984"));

            xrmBrowser.Entity.Save();

        }
        public static void Close()
        {
            xrmBrowser.Dispose();
        }
    }
}
