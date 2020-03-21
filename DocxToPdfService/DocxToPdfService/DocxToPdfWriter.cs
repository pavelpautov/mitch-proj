namespace DocxToPdfService
{
    using System;
    using System.IO;
    //using System.Reflection;
    using Word = Microsoft.Office.Interop.Word;

    public static class DocxToPdfService
    {
        public static void Convert(string sourcePath, string destinationPath)
        {
            Word._Application oWord = new Word.Application
            {
                Visible = false
            };

            // Interop requires objects.
            object oMissing = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = true;
            object oInput = sourcePath;
            object oOutput = destinationPath;
            object oFormat = Word.WdSaveFormat.wdFormatPDF;

            if (!File.Exists(sourcePath))
            {
                return;
            }

            // Load a document into our instance of word.exe
            Word._Document oDoc = oWord.Documents.Open(
                ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

            // Make this document the active document.
            oDoc.Activate();

            // Save this document using Word
            oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

            // Always close Word.exe.
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
        }
    }

    /*
    public static class Library
    {
        static FileSystemWatcher fsw = new FileSystemWatcher();
        static string[] locations =
        {
            @"R:\Location1\directory1\test\",
            @"A:\Location2\directory2\test\"
        };

        public static void ConvertToPDF()
        {
            DateTime today = DateTime.Today;
            string s_today = today.ToString("MMddyyyy");

            foreach (string path in locations)
            {
                string fullPath = Path.Combine(path, s_today);

                fsw.Path = fullPath;
                fsw.Filter = "*.docx";

                fsw.Created += OnCreated;
                fsw.EnableRaisingEvents = true;
            }
        }

        private static void OnCreated(object source, FileSystemEventArgs e)
        {
            FileInfo file = new FileInfo(e.FullPath);
            Convert(file.ToString(), Path.GetDirectoryName(e.FullPath) + Path.DirectorySeparatorChar + "Labels.pdf", WdSaveFormat.wdFormatPDF);
            Directory.EnumerateFiles(Path.GetDirectoryName(e.FullPath) + Path.DirectorySeparatorChar, "*.docx").ToList().ForEach(x => File.Delete(x));
        }

        private static void Dispose()
        {
            fsw.Created -= OnCreated;
            fsw.Dispose();
        }

        public static void Convert(string input, string output, WdSaveFormat format)
        {
            _Application oWord = new Word.Application
            {
                Visible = false
            };

            // Interop requires objects.
            object oMissing = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = true;
            object oInput = input;
            object oOutput = output;
            object oFormat = format;

            // Load a document into our instance of word.exe
            _Document oDoc = oWord.Documents.Open(
                ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

            // Make this document the active document.
            oDoc.Activate();

            // Save this document using Word
            oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                );

            // Always close Word.exe.
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
        }
    }*/
}
