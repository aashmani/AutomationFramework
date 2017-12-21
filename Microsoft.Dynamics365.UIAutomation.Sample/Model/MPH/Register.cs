using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
   public class Register
    {
        public static XrmBrowser xrmBrowser;
        static Register()
        {
            BaseModel.GetDataFromYaml();
        }

        public static void NavigateToClientContract()
        {
            xrmBrowser.Navigation.OpenSubArea("Recruitment", "new client/subcontrator contracts");
            Logs.LogHTML("Navigated to Accounts  Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }

        private static void ClickNew()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.CommandBar.ClickCommand("New");
        }

        public static void CreateNewContract()
        {
            ClickNew();

            var dicRegisterClient = BaseModel.jsonObj.SelectToken("RegisterClient");

            xrmBrowser.Entity.SetValue("mph_ismphtemplate");

            xrmBrowser.Entity.SetValue(new Lookup { Name = "mph_mphcontrwithclientinvid", Value = dicRegisterClient["mph_mphcontrwithclientinvid"].ToString() });
            xrmBrowser.Entity.SetValue("mph_contracttype", dicRegisterClient["mph_contracttype"].ToString());
            xrmBrowser.Entity.SetValue("mph_corporatename", dicRegisterClient["mph_corporatename"].ToString());
            xrmBrowser.Entity.SetValue("mph_corporateregistrationno", dicRegisterClient["mph_corporateregistrationno"].ToString());
            xrmBrowser.Entity.SetValue(new Lookup { Name = "mph_countryofregistrationid", Value = dicRegisterClient["mph_countryofregistrationid"].ToString() });
            xrmBrowser.Entity.SetValue("mph_authorizedrepresentativetitle", dicRegisterClient["mph_authorizedrepresentativetitle"].ToString());
            xrmBrowser.Entity.SetValue("mph_authrepname", dicRegisterClient["mph_authrepname"].ToString());
            xrmBrowser.Entity.SetValue("mph_representativeposition", dicRegisterClient["mph_representativeposition"].ToString());
            xrmBrowser.Entity.SetValue("mph_ofcladdress", dicRegisterClient["mph_ofcladdress"].ToString());
            xrmBrowser.Entity.SetValue(new Lookup { Name = "mph_registeredaddcountryregionid", Value = dicRegisterClient["mph_registeredaddcountryregionid"].ToString() });

            xrmBrowser.Entity.SetValue(new OptionSet { Name = "mph_typeofservices", Value = dicRegisterClient["mph_typeofservices"].ToString() });
            xrmBrowser.Entity.SetValue("mph_businessopscoverage", dicRegisterClient["mph_businessopscoverage"].ToString());
            xrmBrowser.Entity.SetValue("mph_agreementcommdate", DateTime.Parse(dicRegisterClient["mph_agreementcommdate"].ToString()));
            xrmBrowser.Entity.SetValue("mph_agreementenddate", DateTime.Parse(dicRegisterClient["mph_agreementenddate"].ToString()));
            xrmBrowser.Entity.SetValue("mph_clientrepname", dicRegisterClient["mph_clientrepname"].ToString());
            xrmBrowser.Entity.SetValue("mph_clientrepposition", dicRegisterClient["mph_clientrepposition"].ToString());
            xrmBrowser.Entity.SetValue("mph_mph_clientrepcontactno1", dicRegisterClient["mph_mph_clientrepcontactno1"].ToString());
            xrmBrowser.Entity.SetValue("mph_clientrepemail", dicRegisterClient["mph_clientrepemail"].ToString());
            //xrmBrowser.Entity.SetValue("mph_authorizedsignatorytitle", dicRegisterClient["mph_authorizedsignatorytitle"].ToString());
            //xrmBrowser.Entity.SetValue("mph_clientauthsignatory", dicRegisterClient["mph_clientauthsignatory"].ToString());
            //xrmBrowser.Entity.SetValue("mph_authorizedsignatoryposition", dicRegisterClient["mph_authorizedsignatoryposition"].ToString());
            xrmBrowser.Entity.SetValue(new Lookup { Name = "mph_mphrepnameid", Value = dicRegisterClient["mph_mphrepnameid"].ToString() });
            //xrmBrowser.Entity.SetValue("mph_mphrepposition", dicRegisterClient["mph_mphrepposition"].ToString());
            //xrmBrowser.Entity.SetValue("mph_mphrepcontactno", dicRegisterClient["mph_mphrepcontactno"].ToString());
            //xrmBrowser.Entity.SetValue("mph_mphrepemail", dicRegisterClient["mph_mphrepemail"].ToString());

            xrmBrowser.Entity.SetValue("mph_cntctaddress", dicRegisterClient["mph_cntctaddress"].ToString());
            xrmBrowser.Entity.SetValue(new Lookup { Name = "mph_contactaddcountryregionid", Value = dicRegisterClient["mph_contactaddcountryregionid"].ToString() });

            xrmBrowser.Entity.SetValue("mph_fctext", dicRegisterClient["mph_fctext"].ToString());
            //xrmBrowser.Entity.SetValue(new Lookup { Name = "mph_invoicingcurrencyid", Value = dicRegisterClient["mph_invoicingcurrencyid"].ToString() });
            //xrmBrowser.Entity.SetValue(new Lookup { Name = "mph_paymentcurrencyid", Value = dicRegisterClient["mph_paymentcurrencyid"].ToString() });
            //xrmBrowser.Entity.SetValue("mph_paymentterms", dicRegisterClient["mph_paymentterms"].ToString());

            xrmBrowser.Entity.SetValue(new Lookup { Name = "mph_businessdevincharge", Value = dicRegisterClient["mph_businessdevincharge"].ToString() });
            xrmBrowser.Entity.SetValue("mph_preparedbydate", DateTime.Parse(dicRegisterClient["mph_preparedbydate"].ToString()));
            xrmBrowser.Entity.SetValue("mph_businessdevinchgdate", DateTime.Parse(dicRegisterClient["mph_businessdevinchgdate"].ToString()));

            xrmBrowser.CommandBar.ClickCommand("Save & Close");
        }

        private static void CopyFromExsistingRecord()
        {
            xrmBrowser.CommandBar.ClickCommand("Copy From Existing CFAIS");
        }

        public static void SelectFirstRecord()
        {
            xrmBrowser.ThinkTime(1000);
            xrmBrowser.Grid.PopupSelectRecord(0);
            Logs.LogHTML("Selected Account", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
    }
}
