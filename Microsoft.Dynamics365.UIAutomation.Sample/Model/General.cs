using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Utility;
using Newtonsoft.Json.Linq;
using System;
using System.Configuration;
using System.IO;
using System.Security;
using YamlDotNet.Serialization;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    public  class General
    {
      
        public static JObject jsonObj;

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

                jsonObj = JObject.Parse(json);
            }
            catch (Exception ex)
            {
                Logs.LogHTML("Get Data From Yaml Failed : " + ex.Message, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            }
        }

        public static void Login(XrmBrowser xrmBrowser, Uri uri, SecureString username, SecureString password,string testCase)
        {
            Logs.LogHTML(string.Empty, Logs.HTMLSection.TestCase, Logs.TestStatus.NA, testCase);
            xrmBrowser.LoginPage.Login(uri, username, password);
            xrmBrowser.GuidedHelp.CloseGuidedHelp();
            Logs.LogHTML("Logged in Successfully", Logs.HTMLSection.Details, Logs.TestStatus.Pass);
        }
        public static void LogError(string ErrorMsg,string TestCase) {
            Logs.LogHTML(TestCase + " : Error : " + ErrorMsg, Logs.HTMLSection.Details, Logs.TestStatus.Fail);
            Helper.failedScenarios.Add(TestCase);
        }
    }
}
