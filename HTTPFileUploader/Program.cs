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


namespace iServerToServiceNowExchanger
{
    class Program
    {
        public static int MaxVersions = 5;
        public static string filePostFix = ".bak.";
        public const int fileLockedTimeout = 5;
        public static string user;
        public static string pass;

        static void Main(string[] args)
        {
            /////////////////////////////////
            // Valid input cases

            // --help

            // Download to file
            // -d <URL> -f <to file>
            
            // Upload to file
            // -u http://mads.it-kartellet.dk/test_uploading.php -f <from file>

            // Merge files to one
            // -m <file_1> -m <file_n> -f <to fileMerged>

            // Split 1 or more files into 1 pr worksheet. (add worksheet name as postfix)
            // -s <file_1> -s <file_n>

            // Download, rotate and merge files.
            // -dm <URL_1> -dm <URL_n> -f <to fileMerged>

            //
            /////////////////////////////////

            
            Log.SetFileListener("iServerToServiceNowExchanger.log");
            //Log.SetLogLevel
            Log.OSInformation();

            bool show_help = false;
            string UploadURL = null;
            string DownloadURL = null;
            string filepath = null;
            List<string> mergeList = new List<string>();
            List<string> downloadMergeList = new List<string>();
            List<string> splitList = new List<string>();

            var p = new OptionSet() {
                { "u|upload=", "the {URL} to upload to.",                                v => UploadURL=v},
                { "d|download=","the {URL} to download from.",                           v => DownloadURL=v},
                { "f|file=", "the {filepath} to use.",                                   v => filepath=v},
                { "m|merge=", "A Excel workbook to merge",                               v => mergeList.Add(v)},
                { "dm|downloadMerge=", "A Excel workbook to download, rotate and merge", v => downloadMergeList.Add(v)},
                { "s|split=", "A Excel workbook to split",                               v => splitList.Add(v)},
                { "user|username=", "The username to login with",                        v => user=v},
                { "pass|password=", "The password to login with",                        v => pass=v}, //FIXME: Is this safe enough?
                { "h|help",  "show this message and exit",                               v => show_help = v != null },
            };

            
            List<string> arguments;
            try
            {
                arguments = p.Parse(args);
            }
            catch (OptionException e)
            {
                Console.Write("iServerToServiceNowExchanger: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try iServerToServiceNowExchanger --help' for more information.");
                Log.Error("Try iServerToServiceNowExchanger --help' for more information.",e);
                return;
            }

            if (show_help)
            {
                showHelp(p);
                return;
            }

            if (user==null ^ pass==null)
            {
                Console.WriteLine("When using credentials, you have to specify both a username and a password.");
                Log.Error("When using credentials, you have to specify both a username and a password.", new Exception()); //FIXME: Log without exception
                return;
            }

            if (DownloadURL != null && filepath != null)
            {
                Console.WriteLine("Downloading " + DownloadURL + " " + filepath);
                Log.Info("Downloading " + DownloadURL + " " + filepath, new Exception()); //FIXME: Log without exception
                downloadAndRotate(DownloadURL, filepath);
                return;
            }
            
            if (UploadURL != null && filepath != null)
            {
                Console.WriteLine("Uploading " + UploadURL + " " + filepath);
                Log.Info("Uploading " + UploadURL + " " + filepath, new Exception()); //FIXME: Log without exception
                fileUpload(UploadURL, filepath);
                return;
            }
            
            if (mergeList.Count != 0 && filepath != null)
            {
                // Merge files from <mergeList> into <filepath>
                List<Workbook> mergeWorkbookList = new List<Workbook>();
                foreach (string file in mergeList)
                {
                    mergeWorkbookList.Add(Workbook.Load(file));
                }

                Workbook workbookMerged = workbookMerge(mergeWorkbookList);

                String dir = Path.GetDirectoryName(filepath) + @"\";
                String filename = Path.GetFileNameWithoutExtension(filepath);
                String extension = Path.GetExtension(filepath);
                workbookMerged.Save(dir + filename + "_merged" + extension);
                return;
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
                        String extension = Path.GetExtension(filename);
                        workbook.Save(dir + file + "_" + i + "_" + workbook.Worksheets.ElementAt(0).Name + extension);
                    }
                }
                return;
            }

            if (downloadMergeList.Count != 0 && filepath != null)
            {
                string dir = Path.GetDirectoryName(filepath) + @"\";

                int i = 0;
                foreach (string URL in downloadMergeList)
                {
                    i++;
                    Console.WriteLine("Downloading " + URL + " " + dir + i + ".tmp");
                    Log.Info("Downloading " + URL + " " + dir + i + ".tmp", new Exception()); //FIXME: Log without exception
                    if (!downloadAndRotate(URL, dir + i + ".tmp"))
                    {
                        Console.WriteLine("Error downloading "+URL+" files will not be merged.");
                        Log.Error("Error downloading " + URL + " files will not be merged.", new Exception()); //FIXME: Log without exception
                        return;
                    }
                }

                List<Workbook> mergeWorkbookList = new List<Workbook>();
                i=1;
                while (File.Exists(dir + i + ".tmp"))
                {
                    mergeWorkbookList.Add(Workbook.Load(dir + i + ".tmp"));
                    i++;
                }
                Workbook workbookMerged = workbookMerge(mergeWorkbookList);

                String filename = Path.GetFileNameWithoutExtension(filepath);
                String extension = Path.GetExtension(filepath);
                workbookMerged.Save(dir + filename + "_merged" + extension);
                return;
            }

            showHelp(p);

            Console.WriteLine("DONE!");
            Console.ReadKey();
        }

        static void showHelp(OptionSet p)
        {
            Console.WriteLine("Usage: HTTPFileUploader [OPTIONS]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
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

        public static void fileUpload(string address, string filepath)
        {
            using (WebClient client = new WebClient())
            {
                //Add Certificate here if needed.
                //if (user != null && pass != null)
                //{
                //    client.Credentials = new System.Net.NetworkCredential(user, pass);
                //}

                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        if (File.Exists(filepath))
                        {
                            var result = client.UploadFile(new Uri(address), filepath);
                            break;
                        }
                        else
                        {
                            Console.WriteLine("File path invalid: " + filepath);
                            Log.Error("File path invalid: " + filepath, new Exception()); //FIXME: Log without exception
                            break;
                        }
                    }
                    catch (WebException e)
                    {
                        Console.WriteLine("File '" + filepath + "' could not be uploaded to server. Retrying 3 times.");
                        Log.Error("File '" + filepath + "' could not be uploaded to server. Retrying 3 times.", e);
                    }
                    catch (UriFormatException e)
                    {
                        Console.WriteLine("URL invalid: " + address);
                        Log.Error("URL invalid: " + address, e);
                        break;
                    }
                }
            }
        }

        public static bool downloadToFile(string url, string filepath)
        {
            filepath += ".tmp";
            if (File.Exists(filepath) && !removeFile(filepath))
            {
                return false;
            }
            var downloaduri = new Uri(url);

            using (var w = new GZipWebClient())
            {
                if (user != null && pass != null)
                {
                    w.Credentials = new System.Net.NetworkCredential(user, pass);
                }

                try
                {
                    w.DownloadFile(downloaduri, filepath);
                    return true;
                }
                 catch (WebException e)
                {
                    Console.WriteLine("File '" + filepath + "' could not be downloaded from server.");
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

        public static bool renameFile(string oldFilepath, string newFilepath, int timeout = fileLockedTimeout)
        {
            if (waitOnFileNotLocked(oldFilepath, timeout))
            {
                try
                {
                    if (File.Exists(newFilepath))
                    {
                        removeFile(newFilepath); //There shouldn't be a file at newFilePath. If so, it is an error.
                        Console.WriteLine("Removed file which shouldn't be there: " + newFilepath);
                        Log.Info("Removed file which shouldn't be there: " + newFilepath, new Exception()); //FIXME: Log without exception
                    }
                    File.Move(oldFilepath, newFilepath);
                    return true;
                }
                catch (Exception e)
                {
                    Console.WriteLine("Could not move file: " + oldFilepath + " to " + newFilepath);
                    Log.Error("Could not move file: " + oldFilepath + " to " + newFilepath, e);
                    return false;
                }
            }
            else
            {
                Console.WriteLine("Could not move '" + oldFilepath + "' because it is locked");
                Log.Error("Could not move '" + oldFilepath + "' because it is locked", new Exception()); //FIXME: Log without exception
                return false;
            }
        }

        public static bool removeFile(string Filepath)
        {
            if (waitOnFileNotLocked(Filepath))
            {
                try
                {
                    File.Delete(Filepath);
                    return true;
                }
                catch (Exception e)
                {
                    Console.WriteLine("Could not delete file: " + Filepath);
                    Log.Error("Could not delete file: " + Filepath, e);
                    return false;
                }
            }
            else
            {
                Console.WriteLine("Could not delete '" + Filepath + "' because it is locked");
                Log.Error("Could not delete '" + Filepath + "' because it is locked", new Exception()); //FIXME: Log without exception
                return false;
            }
        }

        private static bool downloadAndRotate(string url, string filepath)
        {
            if (downloadToFile(url, filepath))
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
                    Console.WriteLine("Operation aborting because either '" + filepath + "' doesn't exist or '" + filepath + "' and '" + filepath + ".tmp' couldn't be renamed/moved");
                    Log.Error("Operation aborting because either '" + filepath + "' doesn't exist or '" + filepath + "' and '" + filepath + ".tmp' couldn't be renamed/moved", new Exception()); //FIXME: Log without exception
                    return false;
                }
            }
            else
            {
                Console.WriteLine("Operation aborted because download failed. Files won't be rotated.");
                Log.Error("Operation aborted because download failed. Files won't be rotated.", new Exception()); //FIXME: Log without exception
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
                    Console.WriteLine("Critical error occured when renaming/moving backup files. Filenames might now have inconsistencies.");
                    Log.Error("Critical error occured when renaming/moving backup files. Filenames might now have inconsistencies.", new Exception()); //FIXME: Log without exception
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
