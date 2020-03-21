namespace DocxToPdfService
{
    using System.IO;
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
}
