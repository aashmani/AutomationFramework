using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Microsoft.Dynamics365.UIAutomation.Utility
{
    public class Logs
    {
        private static string logPath = Path.GetFullPath(@"..\..\..");

        static Logs()
        {
            if (!Directory.Exists(logPath + @"\Logs"))
                Directory.CreateDirectory(logPath + @"\Logs");
            if (!Directory.Exists(logPath + @"\Logs\Errors"))
                Directory.CreateDirectory(logPath + @"\Logs\Errors");
            if (!Directory.Exists(logPath + @"\Logs\TestCases\"))
                Directory.CreateDirectory(logPath + @"\Logs\TestCases\");
            //if (!Directory.Exists(logPath + @"\Logs\Screenshots\"))
            //    Directory.CreateDirectory(logPath + @"\Logs\Screenshots\");
        }
        /// <summary>
        /// To Log Errors in txt file.
        /// </summary>
        /// <param name="ErrorId"></param>
        /// <param name="TraceEventTypeError"></param>
        /// <param name="CommandErrorEventId"></param>
        /// <param name="CommandError"></param>
        /// <param name="CommandName"></param>
        /// <param name="ExecutionAttempts"></param>
        /// <param name="RetryAttempts"></param>
        /// <param name="FullName"></param>
        /// <param name="Message"></param>
        /// <param name="ImageName"></param>
        public static void Log(string ErrorId, string TraceEventTypeError, string CommandErrorEventId, string CommandError = "", string CommandName = "",
                           string ExecutionAttempts = "", string RetryAttempts = "", string FullName = "", string Message = "", string ImageName = "")
        {


            string file = TraceEventTypeError + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            string filename = logPath + @"\Logs\Errors\" + file;

            if (!File.Exists(filename))
            {
                using (StreamWriter sw = File.CreateText(filename))
                {
                    sw.WriteLine("ErrorId : " + ErrorId);
                    //if (CommandErrorEventId != "")
                    //    sw.WriteLine("CommandErrorEventId : " + CommandErrorEventId);
                    if (CommandError != "")
                        sw.WriteLine("CommandError : " + CommandError);
                    if (CommandName != "")
                        sw.WriteLine("CommandName : " + CommandName);
                    if (ExecutionAttempts != "")
                        sw.WriteLine("ExecutionAttempts : " + ExecutionAttempts);
                    if (RetryAttempts != "")
                        sw.WriteLine("RetryAttempts : " + RetryAttempts);
                    if (FullName != "")
                        sw.WriteLine("FullName : " + FullName);
                    if (Message != "")
                        sw.WriteLine("Message : " + Message);
                    if (ImageName != "")
                        sw.WriteLine("FileTimeStamp : " + ImageName);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filename))
                {
                    sw.WriteLine("************************************");

                    sw.WriteLine("ErrorId : " + ErrorId);
                    //if (CommandErrorEventId != "")
                    //    sw.WriteLine("CommandErrorEventId : " + CommandErrorEventId);
                    if (CommandError != "")
                        sw.WriteLine("CommandError : " + CommandError);
                    if (CommandName != "")
                        sw.WriteLine("CommandName : " + CommandName);
                    if (ExecutionAttempts != "")
                        sw.WriteLine("ExecutionAttempts : " + ExecutionAttempts);
                    if (RetryAttempts != "")
                        sw.WriteLine("RetryAttempts : " + RetryAttempts);
                    if (FullName != "")
                        sw.WriteLine("FullName : " + FullName);
                    if (Message != "")
                        sw.WriteLine("Message : " + Message);
                    if (ImageName != "")
                        sw.WriteLine("FileTimeStamp : " + ImageName);
                }
            }
        }

        /// <summary>
        /// To log Details in HTML
        /// </summary>
        /// <param name="message">Message to Show in test result</param>
        /// <param name="section"></param>
        /// <param name="teststatus"></param>
        /// <param name="TestCase"></param>
        /// <param name="User"></param>
        /// <param name="Browser"></param>
        /// 
        public static void LogHTML(string message, HTMLSection section, TestStatus teststatus = TestStatus.NA, string TestCase = "", string User = "", string Browser = "")
        {

            string file = Helper.htmlLogFileName;
            string filename = logPath + @"\Logs\TestCases\" + file;
            StreamWriter sw = null;
           
            if (!File.Exists(filename))
            {
                sw = File.CreateText(filename);
                sw.WriteLine(" <link rel = 'stylesheet' href ='https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta.2/css/bootstrap.min.css' integrity = 'sha384-PsH8R72JQ3SOdhVi3uxftmaW6Vc51MKb0q5P2rRUpPvrszuE4W1povHYgTpBfshb' crossorigin = 'anonymous' >");
                sw.WriteLine(" <style > p {padding-left: 5%; } </style >");

            }
            else
            {
                sw = File.AppendText(filename);              

            }
            using (sw)
            {
                //sw.WriteLine(" < !DOCTYPE html >");
                //sw.WriteLine("< html >");
                //sw.WriteLine(" < body >");

                if (section == HTMLSection.User)
                {
                    sw.WriteLine("<h4 class='bg - info'>  User :  " + User + "</h4>");
                    sw.WriteLine("<h4 class='text - muted'> Browser : " + Browser + "</h4>");
                }
                else if (section == HTMLSection.TestCase)
                {
                    sw.WriteLine("<h5> TestCase : " + TestCase + "</h5>");
                }
                else if (section == HTMLSection.Details && teststatus == TestStatus.Pass)
                    sw.WriteLine("<p style = 'color: green;'> " + message + "</p>");
                else if (section == HTMLSection.Details && teststatus == TestStatus.Fail)
                    sw.WriteLine("<p style = 'color: red;'> " + Regex.Replace(message, "<.*?>", String.Empty) + "</p> ");

                else if (section == HTMLSection.TestCaseCount)
                {
                    sw.WriteLine("<div style = 'color: green;'><strong> " + message + "</strong></div>");
                }
                else if (section == HTMLSection.TestCaseFail)
                {
                    sw.WriteLine("<div style = 'color: red;'><em> " + message + "</em></div>");
                }
                //To be deleted after migrating old code
                else if (section == HTMLSection.Header)
                {
                    sw.WriteLine("<h1> TestCase : " + TestCase + "</h1>");
                    sw.WriteLine("<h2> User :  " + User + "</h2>");
                    sw.WriteLine("<h3> Browser : " + Browser + "</h3>");
                }
                
            } 
        }
        public enum TestStatus { Pass, Fail, NA };

        public enum HTMLSection { Header, User, TestCase, Details, Images,TestCaseCount,TestCaseFail };
    }
}