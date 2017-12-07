using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System.Reflection;
using System;
using TaskScheduler;
using System.Xml.Serialization;
using System.IO;
using System.Xml;
using Microsoft.Win32.TaskScheduler;


namespace Microsoft.Dynamics365.UIAutomation.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<BrowserEntity> lstBrowse = new List<BrowserEntity>();
        List<RolesConfiguration> lstRole = new List<RolesConfiguration>();
        List<string> lstSenario = new List<string>();
        string hostURL = string.Empty;
        //string password = string.Empty;
        //string username = string.Empty;
        List<RoleUsersConfiguration> lstSelectedRoleUsers = new List<RoleUsersConfiguration>();
        string schBrowser = string.Empty;
        string schUser = string.Empty;
        string schRole=string.Empty;
        List<string> lstSchScenario = new List<string>();
        RoleUsersConfiguration schRoleconfig = new RoleUsersConfiguration();

        public MainWindow()
        {
            InitializeComponent();

            var objBrowser = ConfigurationManager.GetInstance().browsers;
            lstRole = ConfigurationManager.GetInstance().Roles.ToList();
            var objEntities = ConfigurationManager.GetInstance().Entities;
            hostURL = ConfigurationManager.GetInstance().rootURL;
     

            //**********************************************
            #region For Run Test
            //**********************************************
            //cb_Roles.ItemsSource = lstRole;
            foreach (var browser in objBrowser)
            {
                lstBrowse.Add(new BrowserEntity
                {
                    BrowserName = browser.ToString()

                });
            }
            cb_Browser.ItemsSource = lstBrowse;
         

            foreach (var entities in objEntities)
            {
                var entityExpander = new Expander { Header = entities.entityname };
                entityExpander.IsExpanded = true;
                var stkSenario = new StackPanel();
                foreach (var scenario in entities.scenarios)
                {
                    var cbScenario = new CheckBox { Content = scenario, Margin = new Thickness(10, 4, 4, 4) };
                    cbScenario.Checked += CheckBoxSenario_Checked;
                    cbScenario.Unchecked += CheckBoxSenario_UnChecked;
                    stkSenario.Children.Add(cbScenario);
                }
                entityExpander.Content = stkSenario;
                stkExpanderMain.Children.Add(entityExpander);
            }

            foreach(var role in lstRole)
            {
                var roleExpander = new Expander { Header = role.rolename };
                roleExpander.IsExpanded = true;
                var stkRole = new StackPanel();
                foreach(var user in role.users)
                {
                    var cbRole = new CheckBox { Content = user.username, Tag = user.password, Margin = new Thickness(10, 4, 4, 4) };
                    cbRole.Checked += CheckBoxRole_Checked;
                    cbRole.Unchecked+=CheckBoxRole_UnChecked;
                    stkRole.Children.Add(cbRole);
                }
                roleExpander.Content = stkRole;
                stkExpanderRoleMain.Children.Add(roleExpander);
            }
            #endregion

            //**********************************************
            #region For Schedule Test
            //**********************************************
            cb_SchBrowser.ItemsSource = lstBrowse;

            foreach (var entities in objEntities)
            {
                var entitySchExpander = new Expander { Header = entities.entityname };
                entitySchExpander.IsExpanded = true;
                var stkSchSenario = new StackPanel();
                foreach (var scenario in entities.scenarios)
                {
                    var rbSchScenario = new CheckBox { Content = scenario, Margin = new Thickness(10, 4, 4, 4) };
                    rbSchScenario.Checked += rbSchScenario_Checked;
                    rbSchScenario.Unchecked += rbSchScenario_UnChecked;
                    stkSchSenario.Children.Add(rbSchScenario);
                }
                entitySchExpander.Content = stkSchSenario;
                stkExpanderSchMain.Children.Add(entitySchExpander);
            }
            foreach (var role in lstRole)
            {
                var roleSchExpander = new Expander { Header = role.rolename };
                roleSchExpander.IsExpanded = true;
                var stkSchRole = new StackPanel();
                foreach (var user in role.users)
                {
                    var rbSchRole = new RadioButton { Content = user.username, Tag= user.password,  Margin = new Thickness(10, 4, 4, 4) };
                    rbSchRole.GroupName = "Role";
                    rbSchRole.Checked += rbSchRole_Checked;
                    rbSchRole.Unchecked += rbSchRole_UnChecked;
                    stkSchRole.Children.Add(rbSchRole);
                }
                roleSchExpander.Content = stkSchRole;
                stkExpanderSchRoleMain.Children.Add(roleSchExpander);
            }
            #endregion
        }


        //**********************************************
        #region For Run Test
        //**********************************************
        private void CheckBoxSenario_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox chkSenario = sender as CheckBox;
            lstSenario.Add(chkSenario.Content.ToString());
        }
        private void CheckBoxSenario_UnChecked(object sender, RoutedEventArgs e)
        {
            CheckBox chkSenario = sender as CheckBox;
            if (lstSenario.Contains(chkSenario.Content.ToString()))
            {
                lstSenario.Remove(chkSenario.Content.ToString());
            }
        }

        private void CheckBoxRole_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox chkRole = sender as CheckBox;
            RoleUsersConfiguration user = new RoleUsersConfiguration();
            user.username = chkRole.Content.ToString();
            user.password = chkRole.Tag.ToString();
            user.IsChecked = true;
            lstSelectedRoleUsers.Add(user);
        }
        private void CheckBoxRole_UnChecked(object sender, RoutedEventArgs e)
        {
            CheckBox chkRole = sender as CheckBox;
            lstSelectedRoleUsers.RemoveAll(item => item.username == chkRole.Content.ToString());
        }

        private void btnRuntest_Click(object sender, RoutedEventArgs e)
        {
            btnRuntest.IsEnabled = false;
            List<BrowserEntity> lstSelectedBrowser = new List<BrowserEntity>();
            lstSelectedBrowser = lstBrowse.Where(b => b.IsChecked.Equals(true)).ToList();           
            string path = System.AppDomain.CurrentDomain.BaseDirectory + @"Microsoft.Dynamics365.UIAutomation.Sample.dll";
            //foreach(var browswer in lstSelectedBrowser)
            //{

            foreach (var user in lstSelectedRoleUsers)
            {
                //foreach (var user in role)
                //{
                    foreach (var scenario in lstSenario)
                    {
                        Assembly objAssembly = Assembly.LoadFile(path);
                        if (objAssembly != null)
                        {
                            Type type = objAssembly.GetType("Microsoft.Dynamics365.UIAutomation.Sample." + scenario);
                            //type.GetMethod("TestUpdateAccount").Invoke(Activator.CreateInstance(type), null);
                            if (type != null)
                            {
                                object objType = Activator.CreateInstance(type);
                                FieldInfo field = type.GetField("_username", BindingFlags.NonPublic | BindingFlags.Instance);
                                field.SetValue(objType, user.username.ToSecureString());
                                FieldInfo fieldPassword = type.GetField("_password", BindingFlags.NonPublic | BindingFlags.Instance);
                                fieldPassword.SetValue(objType, user.password.ToSecureString());
                                FieldInfo fieldURL = type.GetField("_xrmUri", BindingFlags.NonPublic | BindingFlags.Instance);
                                fieldURL.SetValue(objType, new Uri(hostURL));
                          
                                FieldInfo fieldBrowser = type.GetField("_browser", BindingFlags.NonPublic | BindingFlags.Instance);
                                fieldBrowser.SetValue(objType,  BrowserType.Chrome);  //Update BrowserType When  Browser seclection implemented
                            //fieldBrowser.SetValue(objType, lstSelectedBrowser);


                            MethodInfo[] methodInfos = type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
                                if (methodInfos.Count() > 0)
                                {
                                    methodInfos[0].Invoke(objType, null);
                                }
                            }
                        }

                    }
                //}
            }
            btnRuntest.IsEnabled = true;
            //}
        }

        #endregion

        //*****************************************************
        #region For Schedule Test
        //*****************************************************
        private void rbSchRole_UnChecked(object sender, RoutedEventArgs e)
        {
            
        }

        private void rbSchRole_Checked(object sender, RoutedEventArgs e)
        {
            var rbRole = sender as RadioButton;
            schRoleconfig.username = rbRole.Content.ToString();
            schRoleconfig.password = rbRole.Tag.ToString();
            schRoleconfig.IsChecked = true;
            // schRole = rbRole.Content.ToString();
        }

        private void rbSchScenario_UnChecked(object sender, RoutedEventArgs e)
        {
            CheckBox chkSenario = sender as CheckBox;
            if (lstSchScenario.Contains(chkSenario.Content.ToString()))
            {
                lstSchScenario.Remove(chkSenario.Content.ToString());
            }
        }

        private void rbSchScenario_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox chkSenario = sender as CheckBox;
            lstSchScenario .Add(chkSenario.Content.ToString());
        }
        private void btnScheduleTest_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string taskName = txtTaskName.Text;
                if (string.IsNullOrEmpty(taskName))
                {
                    MessageBox.Show("Please enter task name.");
                    return;
                }
                var xmlScenario = XMLSerializer.SerializeObject(lstSchScenario);
                var schArgument = schBrowser + "|" + xmlScenario + "|" + hostURL + "|" + schRoleconfig.username + "|" + schRoleconfig.password;
                //Triggering scheduler using win32 dll
                using (TaskService ts = new TaskService())
                {
                    // Create a new task
                    //const string taskName = "CRMAutomation";
                    Task t = ts.AddTask(taskName,
                       new TimeTrigger()
                       {
                           StartBoundary = DateTime.Now + TimeSpan.FromHours(1),
                           Enabled = false,
                       },
                       new ExecAction(@"D:\AutomationFramework\Microsoft.Dynamics365.UIAutomation.AutomationScheduler\bin\Debug\Microsoft.Dynamics365.UIAutomation.AutomationScheduler.exe", schArgument, null));
                    // Edit task and re-register if user clicks Ok
                    TaskEditDialog editorForm = new TaskEditDialog();
                    editorForm.Editable = true;
                    editorForm.RegisterTaskOnAccept = true;
                    editorForm.Initialize(t);
                    // ** The four lines above can be replaced by using the full constructor
                    editorForm.ShowDialog();
                }

            }
            catch(Exception ex)
            {

            }
            }

       
        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var rbBrowser = sender as RadioButton;
            var txtBrowser = rbBrowser.Content as TextBlock;
            schBrowser = txtBrowser.Text;
        }
        #endregion
    }

    public static class XMLSerializer
    {
        public static string SerializeObject(this List<string> list)
        {
         
            using (MemoryStream xmlStream = new MemoryStream())
            {
                StringWriter stringWriter = new StringWriter();
                XmlDocument xmlDoc = new XmlDocument();

                XmlTextWriter xmlWriter = new XmlTextWriter(stringWriter);

                XmlSerializer serializer = new XmlSerializer(typeof(List<string>));
                XmlSerializerNamespaces xmlNamespace = new XmlSerializerNamespaces();
                xmlNamespace.Add("","");

                serializer.Serialize(xmlWriter, list,xmlNamespace);

                string xmlResult = stringWriter.ToString();

                xmlDoc.LoadXml(xmlResult);
                if (xmlDoc.FirstChild.NodeType == XmlNodeType.XmlDeclaration)
                    xmlDoc.RemoveChild(xmlDoc.FirstChild);
                return xmlDoc.InnerXml;
            }
          
        }
    }

}
