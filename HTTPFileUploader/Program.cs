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

        static int Main(string[] args)
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

            WebProxy proxy = String.IsNullOrWhiteSpace(appSettings["proxy"]) ? null : new WebProxy(appSettings["proxy"]);

            int logLevel = 0;
            try
            {
                logLevel = Int16.Parse(appSettings["logLevel"]);
            }
            catch (FormatException e)
            {
                Log.Warning("Can't parse the log level", e);
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
                return 255;
            }

            if (show_examples)
            {
                showExamples();
                return 255;
            }

            if (show_help || args.Length==0)
            {
                showHelp(p);
                return 255;
            }

            //Log.OSInformation();

            if (downloadService != null && filepath != null)
            {
                string dir = Path.GetDirectoryName(filepath) + @"\";
                string file = Path.GetFileNameWithoutExtension(filepath); //Optional filename will be used as prefix

                if (downloadService.ToLower().Equals("servicenow"))
                {
                    Log.Info("Download from ServiceNow to {0}", filepath);
                    var workbookMerged = new Workbook();
                    foreach (var keyname in appSettings.AllKeys.Where(x => x.StartsWith("serviceNowDownloadURL")))
                    {
                        var sheetname = keyname.Substring("serviceNowDownloadURL".Length);
                        var downloadpath = dir + file + sheetname + ".xls";
                        downloadAndRotate(appSettings[keyname], downloadpath, proxy);
                        
                        // Add worksheets to the merged workbook
                        foreach(var sheet in Workbook.Load(downloadpath).Worksheets) {
                            sheet.Name = sheetname;// +"_" + sheet.Name;
                            workbookMerged.Worksheets.Add(sheet);
                        }
                    }
                    workbookMerged.Save(dir + file + "Merged" + ".xls");
                }
                else
                {
                    Log.Error("Download service '{0}' not valid!", downloadService);
                    return 255;
                }
            }
            else if (uploadService != null && filepath != null)
            {
                Log.Info("Upload {0} to ServiceNow", filepath);
                if (uploadService.ToLower().Equals("servicenow"))
                {
                    var uploadordernames = appSettings["serviceNowUploadOrder"].Split(',').ToList<String>();
                    var worksheets = Workbook.Load(filepath).Worksheets;
                    if (uploadordernames.Count < worksheets.Count)
                    {
                        Log.Error("Number of serviceNowUploadSheetOrder item is smaller than actual sheets in file: {0} < {1}");
                        return 255;
                    }

                    foreach (var sheet in worksheets)
                    {
                        var workbook = new Workbook();
                        workbook.Worksheets.Add(sheet);
                        var uploadname = uploadordernames[0];
                        uploadordernames.RemoveAt(0);
                        if (appSettings["serviceNowUploadURL" + uploadname] != null)
                        {
                            var tmpfile = Path.GetTempFileName();
                            workbook.Save(tmpfile);
                            if (!fileUpload(appSettings["serviceNowUploadURL" + uploadname], tmpfile, proxy))
                            {
                                return 255;
                            }
                            removeFile(tmpfile);
                        }
                        else
                        {
                            Log.Error("Could not find upload configuration serviceNowUploadURL" + uploadname);
                            return 255;
                        }
                    }                    
                }
                else
                {
                    Log.Error("Upload service '{0}' not valid!", uploadService);
                    return 255;
                }
            }
            return 0;
        }

        static void showHelp(OptionSet p)
        {
            Console.WriteLine("Usage: HTTPFileUploader [OPTIONS]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
            Console.WriteLine("-f can be placed anywhere and will be universal for all the operations.");
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

        public static bool fileUpload(string url, string filepath, WebProxy wp = null)
        {
            using (WebClient client = new WebClient())
            {
                if (wp != null)
                {
                    client.Proxy = wp;
                }

                var uploadurl = new Uri(url);

                // Get the username and password if it has been set
                if (!string.IsNullOrEmpty(uploadurl.UserInfo))
                {
                    var credInfo = uploadurl.UserInfo.Split(':');
                    if (credInfo.Length == 2)
                    {
                        client.Credentials = new System.Net.NetworkCredential(credInfo[0], credInfo[1]);
                    }

                    // Strip userinfo
                    uploadurl = new Uri(uploadurl.GetComponents(UriComponents.AbsoluteUri & ~UriComponents.UserInfo, UriFormat.UriEscaped));
                }

                try
                {
                    if (File.Exists(filepath))
                    {
                        Log.Trace("Uploading {0} {1}", uploadurl.ToString(), filepath);
                        var result = client.UploadFile(uploadurl, filepath);
                        Log.Trace("Uploaded {0} {1}", uploadurl.ToString(), filepath);
                        return true;
                    }
                    else
                    {
                        Log.Error("File path invalid: {0}", filepath);
                        return false;
                    }
                }
                catch (WebException e)
                {
                    Log.Error("File '" + filepath + "' could not be uploaded to server. Retrying 3 times.", e);
                    return false;
                }
                catch (UriFormatException e)
                {
                    Log.Error("URL invalid: " + url, e);
                    return false;
                }
            }
        }

        public static bool downloadToFile(string url, string filepath, WebProxy wp = null)
        {
            filepath += ".tmp";
            if (File.Exists(filepath) && !removeFile(filepath))
            {
                return false;
            }

            Uri downloadurl = new Uri(url);
            using (var client = new GZipWebClient())
            {
                if (wp != null)
                {
                    client.Proxy = wp;
                }

                // Get the username and password if it has been set
                if (!string.IsNullOrEmpty(downloadurl.UserInfo)) {
                    var credInfo = downloadurl.UserInfo.Split(':');
                    if (credInfo.Length == 2)
                    {
                        client.Credentials = new System.Net.NetworkCredential(credInfo[0], credInfo[1]);
                    }

                    // Strip userinfo
                    downloadurl = new Uri(downloadurl.GetComponents(UriComponents.AbsoluteUri & ~UriComponents.UserInfo, UriFormat.UriEscaped));
                }
                
                try
                {
                    Log.Trace("Downloading {0} to {1}", downloadurl.ToString(), filepath);
                    client.DownloadFile(downloadurl, filepath);
                    Log.Debug("Downloaded {0} to {1}", downloadurl.ToString(), filepath);
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
