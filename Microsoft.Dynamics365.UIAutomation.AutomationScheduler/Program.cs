using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace Microsoft.Dynamics365.UIAutomation.AutomationScheduler
{
   public class Program
    {
        public static string LogCallsStratTime;
        static void Main(string[] args)
        {
            string hostURL = string.Empty;
            string password = string.Empty;
            string username = string.Empty;
            try
            {

                string paramList = string.Join("", args);
                //string paramList = "Chrome|testsalesrep@servicesource.com|<ArrayOfString><string>CreateAccount</string><string>UpdateAccount</string></ArrayOfString>|https://pesdemo3.crm.dynamics.com|admin@pesdemo3.onmicrosoft.com|pass@word1";
                string[] paramArray = paramList.Split('|');

                hostURL = paramArray[2].ToString();
                username = paramArray[3].ToString();
                password = paramArray[4].ToString();
                List<string> lstScenario = new List<string>();
                XmlSerializer serializer = new XmlSerializer(typeof(List<string>), new XmlRootAttribute("ArrayOfString"));
                StringReader stringReader = new StringReader(paramArray[1].ToString());
                lstScenario = (List<string>)serializer.Deserialize(stringReader);
                string path = System.AppDomain.CurrentDomain.BaseDirectory + @"Microsoft.Dynamics365.UIAutomation.Sample.dll";
                Assembly objAssembly = Assembly.LoadFile(path);
                foreach (var scenario in lstScenario)
                {
                    if (objAssembly != null)
                    {
                        Type type = objAssembly.GetType("Microsoft.Dynamics365.UIAutomation.Sample." + scenario);
                        if (type != null)
                        {
                            object objType = Activator.CreateInstance(type);
                            FieldInfo field = type.GetField("_username", BindingFlags.NonPublic | BindingFlags.Instance);
                            field.SetValue(objType, username.ToSecureString());
                            FieldInfo fieldPassword = type.GetField("_password", BindingFlags.NonPublic | BindingFlags.Instance);
                            fieldPassword.SetValue(objType, password.ToSecureString());
                            FieldInfo fieldURL = type.GetField("_xrmUri", BindingFlags.NonPublic | BindingFlags.Instance);
                            fieldURL.SetValue(objType, new Uri(hostURL));
                            FieldInfo fieldBrowser = type.GetField("_browser", BindingFlags.NonPublic | BindingFlags.Instance);
                            fieldBrowser.SetValue(objType,  BrowserType.Chrome); 
                            MethodInfo[] methodInfos = type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
                            if (methodInfos.Count() > 0)
                            {
                                methodInfos[0].Invoke(objType, null);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
    }
    public static class StringExtensions
    {
        public static SecureString ToSecureString(this string @string)
        {
            var secureString = new SecureString();

            if (@string.Length > 0)
            {
                foreach (var c in @string.ToCharArray())
                    secureString.AppendChar(c);
            }

            return secureString;
        }
    }
}
