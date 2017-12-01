using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace Microsoft.Dynamics365.UIAutomation.UI
{ 
    public class ConfigurationManager
    {
        private static readonly ConfigurationManager _INSTANCE = new ConfigurationManager();

        public string ApplicationName { get; private set; }

        public string ClientName { get; private set; }

        public RolesConfiguration[] Roles { get; private set; }

        public EntitiesConfiguration[] Entities { get; private set; }

        public String AttachmentServerURL { get; private set; }

        public string[] browsers { get; private set;  }

        public AdminConfiguration AdminConfiguration { get; private set; }
        public string rootURL { get; private set; }

        public string Password { get; private set; }

        public string UserName { get; private set; }

        private ConfigurationManager()
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(CRMAutomationType));
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.IgnoreWhitespace = true;
                settings.ValidationType = ValidationType.Schema;
                settings.Schemas = new XmlSchemaSet();
                
                settings.Schemas.Add("", AppDomain.CurrentDomain.BaseDirectory + "config\\config.xsd");
                settings.ValidationFlags = XmlSchemaValidationFlags.ProcessInlineSchema |
                                           XmlSchemaValidationFlags.ProcessSchemaLocation |
                                           XmlSchemaValidationFlags.ReportValidationWarnings;
               // settings.ValidationEventHandler += new ValidationEventHandler(HandleXMLError);

                XmlReader reader = null;
                CRMAutomationType root = null;

                try
                {
                    //Logger.log("Configuration: " + AppDomain.CurrentDomain.BaseDirectory + "config\\config.xml");
                    reader = XmlReader.Create(AppDomain.CurrentDomain.BaseDirectory + "config\\config.xml", settings);
                    reader.MoveToContent();
                   
                    root = (CRMAutomationType)serializer.Deserialize(reader);

                    //AdminConfiguration = root.adminconfiguration.;
                    browsers = root.adminconfiguration.browsers;
                    Roles = root.crmconfiguration.roles;
                    Entities = root.crmconfiguration.entities;
                    rootURL = root.crmconfiguration.hosturl;
                    Password = root.crmconfiguration.password;
                    UserName = root.crmconfiguration.username;

                }
                finally
                {
                    if (reader != null) reader.Close();
                }

                string ApplicationName = root.applicationname;
            }
            catch (Exception ex)
            {
            }
        }

        public static ConfigurationManager GetInstance() { return _INSTANCE; }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute("mscrm-automation", Namespace = "", IsNullable = false)]
    public partial class CRMAutomationType
    {

        private string applicationnameField;

        private string clientnameField;

        private AdminConfiguration adminconfigurationField;

        private CRMConfiguration crmconfigurationField;


        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("application-name")]
        public string applicationname
        {
            get
            {
                return this.applicationnameField;
            }
            set
            {
                this.applicationnameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("client-name")]
        public string clientname
        {
            get
            {
                return this.clientnameField;
            }
            set
            {
                this.clientnameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("admin-configuration")]
        public AdminConfiguration adminconfiguration
        {
            get
            {
                return this.adminconfigurationField;
            }
            set
            {
                this.adminconfigurationField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("crm-configuration")]
        public CRMConfiguration crmconfiguration
        {
            get
            {
                return this.crmconfigurationField;
            }
            set
            {
                this.crmconfigurationField = value;
            }
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class AdminConfiguration
    {

        private string dbconnectstringField;

        private EmailConfiguration emailconfigurationField;

        private string[] browsersField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("db-connect-string")]
        public string dbconnectstring
        {
            get
            {
                return this.dbconnectstringField;
            }
            set
            {
                this.dbconnectstringField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("email-configuration")]
        public EmailConfiguration emailconfiguration
        {
            get
            {
                return this.emailconfigurationField;
            }
            set
            {
                this.emailconfigurationField = value;
            }
        }

        /// <remarks/>
        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("browser-name", IsNullable = false)]
        public string[] browsers
        {
            get
            {
                return this.browsersField;
            }
            set
            {
                this.browsersField = value;
            }
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EmailConfiguration
    {

        private string smtpserverField;

        private string fromaddressField;

        private string templatelocationField;

        private string imagelocationField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("smtp-server")]
        public string smtpserver
        {
            get
            {
                return this.smtpserverField;
            }
            set
            {
                this.smtpserverField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("from-address")]
        public string fromaddress
        {
            get
            {
                return this.fromaddressField;
            }
            set
            {
                this.fromaddressField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("template-location")]
        public string templatelocation
        {
            get
            {
                return this.templatelocationField;
            }
            set
            {
                this.templatelocationField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("image-location")]
        public string imagelocation
        {
            get
            {
                return this.imagelocationField;
            }
            set
            {
                this.imagelocationField = value;
            }
        }
    }

   /* /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class BrowsersConfiguration
    {

        private string browserName;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("browser-name")]
        public string browsername
        {
            get
            {
                return this.browserName;
            }
            set
            {
                this.browserName = value;
            }
        }
    }*/

    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class CRMConfiguration
    {

        private string hosturlField;

        private string tenantnameField;

        private string usernameField;

        private string userguidField;

        private string passwordField;

        private RolesConfiguration[] rolesField;

        private EntitiesConfiguration[] entitiesField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("host-url")]
        public string hosturl
        {
            get
            {
                return this.hosturlField;
            }
            set
            {
                this.hosturlField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("tenant-name")]
        public string tenantname
        {
            get
            {
                return this.tenantnameField;
            }
            set
            {
                this.tenantnameField = value;
            }
        }

        /// <remarks/>
        public string username
        {
            get
            {
                return this.usernameField;
            }
            set
            {
                this.usernameField = value;
            }
        }

        /// <remarks/>
        public string userguid
        {
            get
            {
                return this.userguidField;
            }
            set
            {
                this.userguidField = value;
            }
        }

        /// <remarks/>
        public string password
        {
            get
            {
                return this.passwordField;
            }
            set
            {
                this.passwordField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("role", IsNullable = false)]
        public RolesConfiguration[] roles
        {
            get
            {
                return this.rolesField;
            }
            set
            {
                this.rolesField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("entity", IsNullable = false)]
        public EntitiesConfiguration[] entities
        {
            get
            {
                return this.entitiesField;
            }
            set
            {
                this.entitiesField = value;
            }
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class RolesConfiguration
    {

        private string rolenameField;

        private List<RoleUsersConfiguration> usersField;
        private bool isChecked;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("role-name")]
        public string rolename
        {
            get
            {
                return this.rolenameField;
            }
            set
            {
                this.rolenameField = value;
            }
        }


        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("users")]
        public List<RoleUsersConfiguration> users
        {
            get
            {
                return this.usersField;
            }
            set
            {
                this.usersField = value;
            }
        }
        public bool IsChecked
        {
            get
            {
                return isChecked;
            }
            set
            {
                isChecked = value;
            }
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class RoleUsersConfiguration
    {

        private string userField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("user")]
        public string user
        {
            get
            {
                return this.userField;
            }
            set
            {
                this.userField = value;
            }
        }
    }

    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EntitiesConfiguration
    {

        private string entitynameField;

        private string[] scenariosField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("entity-name")]
        public string entityname
        {
            get
            {
                return this.entitynameField;
            }
            set
            {
                this.entitynameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlArrayItemAttribute("scenario-name", IsNullable = false)]
        public string[] scenarios
        {
            get
            {
                return this.scenariosField;
            }
            set
            {
                this.scenariosField = value;
            }
        }
    }

}
