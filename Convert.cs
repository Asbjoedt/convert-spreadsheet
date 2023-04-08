using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Diagnostics;

namespace Convert.Spreadsheet
{
    public class Convert
    {
        public bool OOXML(string input_filepath, string output_filepath)
        {
            bool success = false;

			using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input_filepath, true))
			{
				try
				{
					// Check for certain protection
					if (spreadsheet.WorkbookPart.Workbook.WorkbookProtection != null || spreadsheet.WorkbookPart.Workbook.FileSharing != null) // This line will throw NullReferebceException
					{
						throw new FileFormatException();
					}
				}
				catch (System.NullReferenceException)
				{
					throw new FileFormatException();
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

        // Convert using Excel
        public bool AnyFileFormat_Excel(string input_filepath, string output_filepath, int xlfileformat)
        {
            bool success = false;

            // Open Excel with no window prompts and create workbook instance
            Excel.Application app = new Excel.Application();
            app.DisplayAlerts = false;
            Excel.Workbook wb = app.Workbooks.Open(input_filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false);

            // Save as .xlsx Strict
            wb.SaveAs(output_filepath, xlfileformat);

            // Close Excel
            wb.Close();
            app.Quit();
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
            {
                Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                Marshal.ReleaseComObject(app); // Delete Excel task in task manager
            }

            // Inform user of success
            Console.WriteLine("File was successfully converted");

            // Return success
            success = true;
            return success;
        }

        // Convert using LibreOffice
        public bool AnyFileFormat_LibreOffice(string input_filepath, string output_folder, string output_fileformat)
        {
            bool success = false;
            Process app = new Process();

            // If app is run on Windows
            string? dir = null;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                dir = Environment.GetEnvironmentVariable("LibreOffice");
            }
            if (dir != null)
            {
                app.StartInfo.FileName = dir;
            }
            else
            {
                app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
            }

            app.StartInfo.Arguments = $"--headless --convert-to {output_fileformat} {input_filepath} --outdir {output_folder}";
            app.Start();
            app.WaitForExit();
            app.Close();

            success = true;
            return success;
        }
    }
}
