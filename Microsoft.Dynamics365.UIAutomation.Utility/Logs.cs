using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Dynamics365.UIAutomation.Utility
{
    public class Logs
    {
        private static string  logPath=Path.GetFullPath(@"..\..\..");
        public static void Log(string ErrorId, string TraceEventTypeError, string CommandErrorEventId, string CommandError = "", string CommandName = "",
                           string ExecutionAttempts = "", string RetryAttempts = "", string FullName = "", string Message = "", string ImageName="")
        {

          
            string file = TraceEventTypeError + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            string filename = logPath + @"\Logs\Errors\" + file;
            if (!File.Exists(filename))
            {
                using (StreamWriter sw = File.CreateText(filename))
                {
                    sw.WriteLine("ErrorId : " + ErrorId);
                    if (CommandErrorEventId != "")
                        sw.WriteLine("CommandErrorEventId : " + CommandErrorEventId);
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
                    if (CommandErrorEventId != "")
                        sw.WriteLine("CommandErrorEventId : " + CommandErrorEventId);
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

        public static void LogHTML(string testCaseFile, string strComment, HTMLSection section, TestStatus teststatus =TestStatus.NA, string TestCase = "", string User = "", string Browser = "")
        {

            string file = "CRM Testing-" + testCaseFile + ".html";
            string filename = logPath + @"\Logs\TestCases\" + file;
            if (!File.Exists(filename))
            {
                using (StreamWriter sw = File.CreateText(filename))
                {
                    //sw.WriteLine(" < !DOCTYPE html >");
                    //sw.WriteLine("< html >");
                    //sw.WriteLine(" < body >");


                    if (section == HTMLSection.Header) {
                        sw.WriteLine("<h1> TestCase : " + strComment + "</h1> ");
                            sw.WriteLine("<h2> User :  " + User + "</h2> ");
                            sw.WriteLine("<h3> Browser : " + Browser + "</h3> ");
                    }
                    else if (section == HTMLSection.Details && teststatus == TestStatus.Pass)
                        sw.WriteLine("<p style = 'color: green;'> " + strComment + "</p> ");
                    else if (section == HTMLSection.Details && teststatus == TestStatus.Fail)
                        sw.WriteLine("<p style = 'color: red; '> " + strComment + "</p> ");
                  
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filename))
                {
                   
                    if (section == HTMLSection.Details && teststatus == TestStatus.Pass)
                        sw.WriteLine("<p style = 'color: green;'> " + strComment + "</p> ");
                    else if (section == HTMLSection.Details && teststatus == TestStatus.Fail)
                        sw.WriteLine("<p style = 'color: red; '> " + strComment + "</p> ");

                    //else if (section == HTMLSection.Close)
                    //{
                    //    sw.WriteLine("  </ body >");
                    //    sw.WriteLine(" </ html >");
                    //}
                                      
                }
            }
        }
    
        public  enum TestStatus { Pass, Fail, NA  };
        public enum HTMLSection {  Header, Details, Images, Close };
    }
}