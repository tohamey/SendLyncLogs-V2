using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Extensibility;

namespace ConsoleApplication1
{
    class Program
    {
        //constracting the location of the log file to be zipped
        static string path1 = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        //below is path for logs files used by Office 16.0 (SkypeFB) change the 16.0 to 15.0 for (lync)
        static string path2_16 = @"AppData\Local\Microsoft\Office\16.0\Lync\Tracing\";
        static string path2_15 = @"AppData\Local\Microsoft\Office\15.0\Lync\Tracing\";
        static string _logLocation_16 = Path.Combine(path1, path2_16);
        static string _logLocation_15 = Path.Combine(path1, path2_15);
        //location where i want to save the zipped file
        static string _folderToZip = @"c:\tempSkype4b\logs";
        static string _zippedLogs = @"c:\tempSkype4b\Skype4b_logs.zip";
        static string _lyncLogCopy = @"c:\tempSkype4b\logs\lynclogfile.uccapilog";

        static void Main(string[] args)
        {
            //check if the log file 15 exist, meaning Lync log files
            if (Directory.Exists(_logLocation_15))
            {
                //Array to hold the .uccapilog files
                string[] getLyncLogs = Directory.GetFiles(_logLocation_15, "*.uccapilog");
                //check the temp folder exist and clean
                CreateAndClean(getLyncLogs);

            }
            //check if the log file 16 exist, meaning Skype for business log files
            else if (Directory.Exists(_logLocation_16))
            {
                //Array to hold the .uccapilog files
                string[] getSkypelog = Directory.GetFiles(_logLocation_16, "*.uccapilog");
                //check if the temp folder exist and clean
                CreateAndClean(getSkypelog);

            }
            /*
            following part was written by Christoph Weste and all credit goes to him
            twitter account @_cweste
            */
            string s2 = null;
            if (args.Length > 0)
            {
                Console.WriteLine("User: {0}", args[0]);
            }
            if (args.Length > 1)
            {
                string s = args[1].ToString().Split(':')[1];
                int i = s.Length;
                s2 = s.Substring(0, i - 1);
                Console.WriteLine("Contact: {0}", s2);
                Console.WriteLine("Contact: {0}", args[1]);
            }

            try
            {
                // Create the major UI Automation objects.
                Automation _Automation = LyncClient.GetAutomation();

                // Create a dictionary object to contain AutomationModalitySettings data pairs.
                Dictionary<AutomationModalitySettings, object> _ModalitySettings = new Dictionary<AutomationModalitySettings, object>();

                AutomationModalities _ChosenMode = AutomationModalities.FileTransfer | AutomationModalities.InstantMessage;
                //AutomationModalities _ChosenMode =  AutomationModalities.InstantMessage| AutomationModalities.FileTransfer;

                // Store the file path as an object using the generic List class.
                string myFileTransferPath = string.Empty;
                // Edit this to provide a valid file path.
                myFileTransferPath = @"c:\\tempSkype4b\\Skype4b_logs.zip";

                // Create a generic List object to contain a contact URI.

                String[] invitees = { s2 };
                // Adds text to toast and local user IMWindow text entry control.
                _ModalitySettings.Add(AutomationModalitySettings.FirstInstantMessage, "Hello attached you will get my Skype4B logfile");
                //_ModalitySettings.Add(AutomationModalitySettings.);
                _ModalitySettings.Add(AutomationModalitySettings.SendFirstInstantMessageImmediately, true);

                // Add file transfer conversation context type
                _ModalitySettings.Add(AutomationModalitySettings.FilePathToTransfer, myFileTransferPath);

                // Start the conversation.

                if (invitees != null)
                {
                    IAsyncResult ar = _Automation.BeginStartConversation(
                        _ChosenMode
                        , invitees
                        , _ModalitySettings
                        , null
                        , null);

                    // Block UI thread until conversation is started.
                    _Automation.EndStartConversation(ar);
                }
                //Console.ReadLine();
            }
            catch
            {
                Console.WriteLine("Error");
                Console.ReadLine();
            }
        }
        //Helper Method to create and clean the temp folder
        static void CreateAndClean(String[] args)
        {
            //check if the temp folder exist
            if (Directory.Exists(@"c:\tempSkype4b\logs"))
            {
                //if temp folder exist, check if old logs exist there, if yes delete them
                if (File.Exists(@"c:\tempSkype4b\Skype4b_logs.zip"))
                {
                    //delete the file and use Task to wait for the deletion to finish before start copying
                    Task DeleteFile = Task.Factory.StartNew(() => File.Delete(_zippedLogs));
                    DeleteFile.Wait();
                    //copy and zip the files afterwards
                    CopyAndZip(args);
                }
                //else if folder exist and empty then print to the console that all ready to go
                else
                {
                    CopyAndZip(args);
                }
            }
            //else if the temp folder does not exist, create it
            else
            {
                //wait for creating the directory before starting the zipping process
                Task CreateDirecotry = Task.Factory.StartNew(() => Directory.CreateDirectory(@"c:\\tempSkype4b\\logs"));
                CreateDirecotry.Wait();
                //then copy and zip the files
                CopyAndZip(args);
            }
        }
        //Helper Method to Copy and zip the logs files
        static void CopyAndZip(string[] args)
        {
            foreach (string file in args)
            {
                Task CopyLogFile = Task.Factory.StartNew(() => File.Copy(file, _lyncLogCopy, true));
                CopyLogFile.Wait();
                //zip the copied log file
                ZipFile.CreateFromDirectory(_folderToZip, _zippedLogs, CompressionLevel.Optimal, false);
            }
        }
    }
}
