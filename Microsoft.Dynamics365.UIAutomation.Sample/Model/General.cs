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
    public class General
    {
        static Random rnd = new Random();
        static XrmBrowser xrmBrowser = new XrmBrowser(TestSettings.Options);
        static Dictionary<string, object> dictionaryMain = new Dictionary<string, object>();
        public static Dictionary<string, string> dictionaryCreateAccount = new Dictionary<string, string>();
        public static Dictionary<string, string> dictionaryUpdateAccount = new Dictionary<string, string>();

        static General()
        {
        }

        public static void GetDataFromYaml()
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
            catch (Exception ex)
            {

            }
        }
    }
}
