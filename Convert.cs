using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace convert_spreadsheet
{
    public class Convert
    {
        public bool Convert_OOXML(string input_filepath, string output_filepath)
        {
            bool success = false;

            // If password-protected or reserved by another user
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input_filepath, false))
            {
                if (spreadsheet.WorkbookPart.Workbook.WorkbookProtection != null || spreadsheet.WorkbookPart.Workbook.FileSharing != null)
                {
                    Console.WriteLine("File failed conversion because of write-protection");
                    return success;
                }
            }

            // Convert spreadsheet
            byte[] byteArray = File.ReadAllBytes(input_filepath);
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, true))
                {
                    spreadsheet.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                }
                File.WriteAllBytes(output_filepath, stream.ToArray());
            }

            // Repair spreadsheet
            Repair rep = new Repair();
            rep.Repair_OOXML(output_filepath);

            // Inform user of success
            Console.WriteLine("File was successfully converted");

            // Return success
            success = true;
            return success;
        }

        public bool Convert_All(string input_filepath, string output_filepath)
        {
            bool success = false;

            // Open Excel with no window prompts and create workbook instance
            Excel.Application app = new Excel.Application();
            app.DisplayAlerts = false;
            Excel.Workbook wb = app.Workbooks.Open(input_filepath);

            // Save as .xlsx Strict
            wb.SaveAs(output_filepath, 61); // 61 is code for Open XML Strict in Excel

            // Close Excel
            wb.Close();
            app.Quit();
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
            {
                Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                Marshal.ReleaseComObject(app); // Delete Excel task in task manager
            }

            // Inform user of success
            Console.WriteLine("File was successfully converted to .xlsx Strict conformance");

            Repair rep = new Repair();
            rep.Repair_OOXML(output_filepath);

            // Return success
            success = true;
            return success;
        }
    }
}
