using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using NDesk.Options;
using ExcelLibrary.SpreadSheet;
using SimpleLog;
using System.Configuration;
using System.Collections.Specialized;

//NOTE: Cell widths/height are not kept when copying!
//NOTE: Workbook-sheets are merged in the listed order.

namespace iServerToServiceNowExchanger
{
    class Program
    {
        public static int MaxVersions = 5;
        public static string filePostFix = ".bak.";
        public const int fileLockedTimeout = 5;
        public static string user;
        public static string pass;

        private const string STD_SERVICENOW = "ServiceNow";
        private const string STD_OBJECT = "_ObjectsTable";
        private const string STD_RELATION = "_RelationsTable";
        private const string STD_MERGED = "_Merged";
        private const string STD_ISERVER = "iServer";
        private const string STD_ISERVER_SPLITTED = "_Splitted";

        static void Main(string[] args)
        {
            /////////////////////////////////
            // Valid input cases

            // --help

            // Download to file
            // -d <servicenow/iserver> -f <to file>
            
            // Upload to file
            // -u <servicenow/iserver> -f <from file>

            // Merge files to one
            // -m <file_1> -m <file_n> -f <to fileMerged>

            // Split 1 or more files into 1 pr worksheet. (adds worksheet name as postfix automaticaly)
            // -s <file_1> -s <file_n>

            //
            /////////////////////////////////

            /////////////////////////////////
            // Loading configuration
            NameValueCollection appSettings = ConfigurationManager.AppSettings;
            Log.SetFileListener(appSettings["logFile"]);
            Log.SetConsoleListener(true);
            user = appSettings["username"];
            pass = appSettings["password"];

            WebProxy proxy = String.IsNullOrWhiteSpace(appSettings["proxy"]) ? null : new WebProxy(appSettings["proxy"]);

            int logLevel = 0;
            try
            {
                logLevel = Int16.Parse(appSettings["logLevel"]);
            }
            catch (FormatException e)
            {
                Log.Error("Can't parse the log level",e);
                // logLevel defaults to 0
            }
            
            switch (logLevel)
            {
                case 0:
                    Log.SetLogLevel(SimpleLog.LogLevel.ERROR);
                    break;
                case 1:
                    Log.SetLogLevel(SimpleLog.LogLevel.WARN);
                    break;
                case 2:
                    Log.SetLogLevel(SimpleLog.LogLevel.INFO);
                    break;
                case 3:
                    Log.SetLogLevel(SimpleLog.LogLevel.DEBUG);
                    break;
                case 4:
                    Log.SetLogLevel(SimpleLog.LogLevel.TRACE);
                    break;
                default:
                    Log.SetLogLevel(SimpleLog.LogLevel.ERROR);
                    break;
            }

            //
            /////////////////////////////////

            bool show_help = false;
            bool show_examples = false;
            string uploadService = null;
            string downloadService = null;
            string filepath = null;
            List<string> mergeList = new List<string>();
            List<string> splitList = new List<string>();

            var p = new OptionSet() {
                { "f|file=", "the {filepath} to use.",                                     v => filepath = v},
                { "u|upload=", "the {service} to upload to.",                              v => uploadService = v},
                { "d|download=","the {service} to download from.",                         v => downloadService = v},
                { "m|merge=", "A {Excel workbook} to merge",                               v => mergeList.Add(v)},
                { "s|split=", "A {Excel workbook} to split",                               v => splitList.Add(v)},
                { "h|help",  "show this message",                                          v => show_help = v != null },
                { "e|example",  "show examples",                                           v => show_examples = v != null }
            };
            
            List<string> arguments;
            try
            {
                arguments = p.Parse(args);
            }
            catch (OptionException e)
            {
                Log.Error("Try iServerToServiceNowExchanger --help for more information.", e);
                return;
            }

            if (show_examples)
            {
                showExamples();
                return;
            }

            if (show_help || args.Length==0)
            {
                showHelp(p);
                return;
            }

            Log.OSInformation();

            if (user==null ^ pass==null)
            {
                Log.Error("You have to specify both a username and a password.");
                return;
            }

            if (downloadService != null && filepath != null)
            {
                string dir = Path.GetDirectoryName(filepath) + @"\";
                string file = Path.GetFileNameWithoutExtension(filepath); //Optional filename will be used as prefix

                if (downloadService.ToLower().Equals("servicenow"))
                {
                    if (downloadAndRotate(appSettings["serviceNowDownloadURLObject"], dir + file + STD_OBJECT + ".xls", proxy) &&
                        downloadAndRotate(appSettings["serviceNowDownloadURLRelation"], dir + file + STD_RELATION + ".xls", proxy))
                    {
                        List<Workbook> mergeWorkbookList = new List<Workbook>();
                        mergeWorkbookList.Add(Workbook.Load(dir + file + STD_RELATION + ".xls"));
                        mergeWorkbookList.Add(Workbook.Load(dir + file + STD_OBJECT + ".xls"));
                        Workbook workbookMerged = workbookMerge(mergeWorkbookList);
                        workbookMerged.Save(dir + file + STD_MERGED + ".xls");
                    }
                }
                else if (downloadService.ToLower().Equals("iserver"))
                {
                    if (downloadAndRotate(appSettings["iServerDownloadURL"], dir + file + ".xls", proxy))
                    {
                        List<Workbook> workbookSheets = workbookSplit(Workbook.Load(dir + file + ".xls"));
                        int i = 0;
                        foreach (Workbook workbook in workbookSheets)
                        {
                            i++;
                            String extension = Path.GetExtension(filepath);
                            string postfix;
                            switch (i)
                            {
                                case 1: postfix = STD_RELATION; break;
                                case 2: postfix = STD_OBJECT; break;
                                default: postfix = i.ToString(); break;
                            }
                            workbook.Save(dir + file + postfix + ".xls");
                        }
                    }
                }
                else
                {
                    Log.Error("Download service '{0}' not valid!", downloadService);
                    return;
                }
            }
            
            if (uploadService != null && filepath != null)
            {
                string dir = Path.GetDirectoryName(filepath) + @"\";
                string file = Path.GetFileNameWithoutExtension(filepath); //Optional filename will be used as prefix

                if (uploadService.ToLower().Equals("servicenow"))
                {
                    fileUpload(appSettings["serviceNowUploadURLRelation"], dir + file + STD_RELATION + ".xls");
                    fileUpload(appSettings["serviceNowUploadURLObject"], dir + file + STD_OBJECT + ".xls");
                }
                else if (uploadService.ToLower().Equals("iserver"))
                {
                    fileUpload(appSettings["iServerUploadURL"], dir + file + ".xls");
                }
                else
                {
                    Log.Error("Upload service '{0}' not valid!", uploadService);
                    return;
                }
            }

            if (mergeList.Count >= 1 && filepath != null)
            {
                // Merge files from <mergeList> into <filepath>
                List<Workbook> mergeWorkbookList = new List<Workbook>();
                foreach (string file in mergeList)
                {
                    //try
                    //{
                        mergeWorkbookList.Add(Workbook.Load(file));
                    //}
                    //catch (){

                    //}
                }

                Workbook workbookMerged = workbookMerge(mergeWorkbookList);
                workbookMerged.Save(filepath);
            }

            if (splitList.Count != 0)
            {
                // Split files from <splitList> and give new files a postfix of: "_<index of sheet>_<name of sheet>"
                foreach (string filename in splitList)
                {
                    List<Workbook> workbookSheets = workbookSplit(Workbook.Load(filename));

                    int i = 0;
                    foreach (Workbook workbook in workbookSheets)
                    {
                        i++;
                        String dir = Path.GetDirectoryName(filename) + @"\";
                        String file = Path.GetFileNameWithoutExtension(filename);
                        String extension = Path.GetExtension(filepath);
                        string postfix;
                        switch (i)
                        {
                            case 1: postfix = STD_RELATION; break;
                            case 2: postfix = STD_OBJECT; break;
                            default: postfix = i.ToString(); break;
                        }
                        workbook.Save(dir + file + postfix + ".xls");
                    }
                }
            }
        }

        static void showHelp(OptionSet p)
        {
            Console.WriteLine("Usage: HTTPFileUploader [OPTIONS]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
            Console.WriteLine("\nYou can chain the operations in this order: -d -u -m -s");
            Console.WriteLine("-f can be placed anywhere and will be universal for all the operations.");
            Console.WriteLine("-m requires 2 or more paths to work.");
        }

        static void showExamples()
        {
            Console.WriteLine(); 
            Console.WriteLine(" Download file from service");
            Console.WriteLine(" -d <servicenow/iserver> -f <to file>");
            Console.WriteLine();
            Console.WriteLine(" Upload file to service");
            Console.WriteLine(" -u <servicenow/iserver> -f <from file>");
            Console.WriteLine();
            Console.WriteLine(" Merge 2 or more Excel workbooks into one");
            Console.WriteLine(" -m <file_1> -m <file_n> -f <to fileMerged>");
            Console.WriteLine();
            Console.WriteLine(" Split 1 or more Excel workbooks into 1 pr worksheet.");
            Console.WriteLine(" -s <file_1> -s <file_n>");
        }

        static List<Workbook> workbookSplit(Workbook workbook)
        {
            List<Workbook> workbookList = new List<Workbook>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var workbookSheet = new Workbook();
                workbookSheet.Worksheets.Add(sheet);
                workbookList.Add(workbookSheet);
            }

            return workbookList;
        }

        static Workbook workbookMerge(List<Workbook> workbookList)
        {
            Workbook workbookMerged = new Workbook();

            //If two sheets have same name, a "_2 ... _n" postfix will be added.
            foreach (Workbook workbook in workbookList)
            {
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    workbookMerged.Worksheets.Add(sheet);
                }
            }

            return workbookMerged;
        }

        public static void fileUpload(string address, string filepath, WebProxy wp = null)
        {
            Log.Trace("Uploading {0} {1}", address, filepath);
            using (WebClient client = new WebClient())
            {
                if (wp != null)
                {
                    client.Proxy = wp;
                }

                if (user != null && pass != null)
                {
                    client.Credentials = new System.Net.NetworkCredential(user, pass);
                }

                try
                {
                    if (File.Exists(filepath))
                    {
                        var result = client.UploadFile(new Uri(address), filepath);
                        return;
                    }
                    else
                    {
                        Log.Error("File path invalid: {0}",filepath);
                        return;
                    }
                }
                catch (WebException e)
                {
                    Log.Error("File '" + filepath + "' could not be uploaded to server. Retrying 3 times.", e);
                }
                catch (UriFormatException e)
                {
                    Log.Error("URL invalid: " + address, e);
                }
            }
        }

        public static bool downloadToFile(string url, string filepath, WebProxy wp = null)
        {
            Log.Trace("Downloading {0} to {1}", url, filepath);
            filepath += ".tmp";
            if (File.Exists(filepath) && !removeFile(filepath))
            {
                return false;
            }

            Uri downloaduri;
            try
            {
                downloaduri = new Uri(url);
            }
            catch (UriFormatException e)
            {
                Log.Error("Check URL address", e);
                return false;
            }

            using (var client = new GZipWebClient())
            {
                if (wp != null)
                {
                    client.Proxy = wp;
                }

                if (user != null && pass != null)
                {
                    client.Credentials = new System.Net.NetworkCredential(user, pass);
                }

                try
                {
                    client.DownloadFile(downloaduri, filepath);
                    Log.Trace("Downloaded {0} to {1}", url, filepath);
                    return true;
                }
                 catch (WebException e)
                {
                    Log.Error("File '" + filepath + "' could not be downloaded from server.", e);
                    return false;
                }
            }
        }

        public class GZipWebClient : WebClient
        {
            protected override WebRequest GetWebRequest(Uri address)
            {
                HttpWebRequest request = (HttpWebRequest)base.GetWebRequest(address);
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                return request;
            }
        }

        public static bool renameFile(string oldfilepath, string newfilepath, int timeout = fileLockedTimeout)
        {
            if (waitOnFileNotLocked(oldfilepath, timeout))
            {
                try
                {
                    if (File.Exists(newfilepath))
                    {
                        removeFile(newfilepath); //There shouldn't be a file at newfilepath. If so, it is an error.
                        Log.Info("Removed file which shouldn't be there: {0}", newfilepath);
                    }
                    File.Move(oldfilepath, newfilepath);
                    return true;
                }
                catch (IOException ex)
                {
                    Log.Error("Could not move file: {0} to {1} : {2}", oldfilepath, newfilepath, ex.Message);
                    return false;
                }
            }
            else
            {
                Log.Error("Could not move '{0}' because it is locked", oldfilepath);
                return false;
            }
        }

        public static bool removeFile(string filepath)
        {
            if (waitOnFileNotLocked(filepath))
            {
                try
                {
                    File.Delete(filepath);
                    return true;
                }
                catch (Exception e)
                {
                    Log.Error("Could not delete file: " + filepath, e);
                    return false;
                }
            }
            else
            {
                Log.Error("Could not delete '{0}' because it is locked", filepath);
                return false;
            }
        }

        private static bool downloadAndRotate(string url, string filepath, WebProxy wp = null)
        {
            if (downloadToFile(url, filepath, wp))
            {
                bool ready = true;
                if (File.Exists(filepath))
                {
                    ready = renameFile(filepath, filepath + filePostFix + 0);
                }

                if (ready && renameFile(filepath + ".tmp", filepath))
                {
                    rotateOldFiles(filepath);
                    return true;
                }
                else
                {
                    Log.Error("Operation aborting because either '{0}' doesn't exist or '{0}' and '{0}.tmp' couldn't be renamed/moved", filepath);
                    return false;
                }
            }
            else
            {
                Log.Error("Operation aborted because download failed. Files won't be rotated.");
                return false;
            }
        }

        private static void rotateOldFiles(string filepath)
        {
            short highest = -1;

            for (int i = 0; i <= MaxVersions; i++) // find highest currently rotated file that isn't higher than MaxRotateCount
            {
                string rotate = filepath + filePostFix + i;
                if (File.Exists(rotate))
                {
                    highest = (short)(i);
                    continue;
                }
                break;
            }

            if (highest == MaxVersions - 1)
            {
                File.Delete(filepath + filePostFix + highest);
                highest--;
            }

            for (; 0 <= highest; highest--) // rotate files backwards
            {
                try
                {
                    if (!renameFile(filepath + filePostFix + highest,
                               filepath + filePostFix + (highest + 1),
                               30))
                    {
                        throw new Exception();
                    }
                }
                catch (Exception)
                {
                    //FIXME: Needs appropriate exception handling. How?
                    Log.Error("Critical error occured when renaming/moving backup files. Filenames might now have inconsistencies.");
                    break;
                }
            }
        }

        public static bool waitOnFileNotLocked(string filename, int timeout = fileLockedTimeout)
        {
            int passed = 0;
            while (isFileLocked(filename))
            {
                Thread.Sleep(1000);
                if (passed++ > timeout)
                {
                    return false;
                }
            }

            return true;
        }

        public static bool isFileLocked(string filename)
        {
            FileInfo file = new FileInfo(filename);
            FileStream stream = null;
            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }
    }
}
