using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace OfficeToHtml
{
    public class XlsxToHtml
    {
        public static void Convert(string absoluteFilePath)
        {
            // The destination file to save must have the extension .html
            string fileToSave = "D:/Output/convertedFile.html"
            Application excel = new Application();
            try
            {
                excel.Workbooks.Open(Filename: absoluteFilePath);
                excel.Visible = false;
                if (excel.Workbooks.Count > 0)
                {
                    IEnumerator wsEnumerator = excel.ActiveWorkbook.Worksheets.GetEnumerator();
                    object format = XlFileFormat.xlHtml;
                    int i = 1;
                    while (wsEnumerator.MoveNext())
                    {
                        var wsCurrent = (Worksheet)wsEnumerator.Current;
                        String outputFile = "excelFile" + "." + i.ToString() + ".html";
                        wsCurrent.SaveAs(Filename: fileToSave, FileFormat: format);
                        ++i;
                        break;
                    }
                    excel.Workbooks.Close();
                }
            }
            finally
            {
                excel.Application.Quit();
            }
        }
    }
}
