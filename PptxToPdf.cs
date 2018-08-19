using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace OfficeToHtml
{
    public class PptxToHtml
    {
        public static string Convert(string absoluteFilePath)
        {
            // The destination file to save must have the extension .html
            string fileToSave = @"D:\Output\convertedFile.html"
            try
            {
                Powerpoint.Presentations.Open(absoluteFilePath, MsoTriState.msoTrue
                , MsoTriState.msoFalse, MsoTriState.msoFalse);
                if (Powerpoint.Presentations.Count > 0)
                {
                    var pptx = Powerpoint.ActivePresentation;
                    pptx.SaveAs(FileName: fileToSave, FileFormat: PpSaveAsFileType.ppSaveAsPDF);
                    pptx.Close();
                }
            }
            finally
            {
                Powerpoint.Quit();
            }
        }
    }
}
