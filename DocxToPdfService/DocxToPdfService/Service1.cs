namespace DocxToPdfService
{
    using System;
    using System.IO;
    using System.ServiceProcess;
    using System.Threading;

    public partial class Service1 : ServiceBase
    {
        public const string DocxFolderPath = @"C:\Pavel\UpWork\Util-for-convert-docx-to-pdf\DocxFolder";
        public const string PdfForderPath  = @"C:\Pavel\UpWork\Util-for-convert-docx-to-pdf\PdfFolder";

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            FileSystemWatcher watcher = new FileSystemWatcher();
            watcher.Path = DocxFolderPath;
            watcher.NotifyFilter = NotifyFilters.LastWrite;
            watcher.Filter = "*.docx";
            watcher.Changed += new FileSystemEventHandler(OnChanged);
            watcher.EnableRaisingEvents = true;
        }

        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            string source = e.FullPath;
            string destination = GeneratePathForDestinationFile(e.FullPath);

            if (!File.Exists(destination))
            {
                WriteToFile($"{DateTime.Now}  Detected file: {Path.GetFileName(e.FullPath)} - event:{e.ChangeType}");
                
                try
                {
                    DocxToPdfService.Convert(source, destination);
                    DeleteFile(source);
                }
                catch (Exception ex)
                {
                    WriteToFile($"exception: {ex} + inner: {ex.InnerException}");
                }
            }
        }

        private static string GeneratePathForDestinationFile(string sourceFilePath)
        {
            return Path.Combine(
                PdfForderPath,
                Path.GetFileNameWithoutExtension(sourceFilePath) + ".pdf");
        }

        private static void DeleteFile(string path)
        {
            int attempts = 10;
            while (true)
            {
                try
                {
                    File.Delete(path);
                    return;
                }
                catch (IOException ioEx)
                {
                    Thread.Sleep(3000);
                    attempts--;
                    if (attempts < 1) return;
                }
            }
        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }

        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
    }
}