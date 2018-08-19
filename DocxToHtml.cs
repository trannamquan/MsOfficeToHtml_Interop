using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace OfficeToHtml
{
    public class DocxToHtml
    {
        public static void Convert(string absoluteFilePath)
        {
            // The destination file to save must have the extension .html
            string fileToSave = "D:/Output/absoluteFileName.html"
            Application Word = new Application();
            try
            {
                Word.Documents.Open(FileName: absoluteFilePath);
                Word.Visible = false;
                if (Word.Documents.Count > 0)
                {
                    Microsoft.Office.Interop.Word.Document oDoc = Word.ActiveDocument;
                    oDoc.SaveAs(FileName: fileToSave, FileFormat: 10);
                    oDoc.Close(SaveChanges: false);
                }
            }
            finally
            {
                Word.Application.Quit(SaveChanges: false);
            }
        }

    }
}
