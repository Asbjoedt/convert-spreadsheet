using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    public static string new_filepath = "";

    string Main(string[] args)
    {
        string filepath = args[0];
        string extension = Path.GetExtension(filepath);
        string message = "";

        try
        {
            switch (extension) // The switch includes all accepted file extensions
            {
                case ".ods":
                case ".ots":
                case ".xls":
                case ".xlsb":
                case ".xlsm":
                case ".xlsx":
                case ".xlt":
                case ".xltm":
                case ".xltx":
                    Console.WriteLine(filepath); // Write filepath to user
                    bool success = Convert_Requirements(filepath); // Convert data
                    if (success == true) // If convert was succesful, write sucess to user
                    {
                        message = "File was successfully converted";
                    }
                    if (filepath != new_filepath) // Delete original file, if it was not a .xlsx
                    {
                        File.Delete(filepath);
                    }
                    return message;

                default:
                    message = "File format is not an accepted file format"; // If the filepath has extension not included in switch
                    return message;
            }
        }
        catch (FileNotFoundException) // If filepath has not file
        {
            message = "No file in filepath"; // Write user of status
            return message;
        }
        catch (FormatException) // If spreadsheet is password protected or otherwise unreadable
        {
            message = "File is password-protected or corrupt"; // Write user of status
            File.Delete(filepath); // Delete file
            return message;
        }
    }

    static bool Convert_Requirements(string filepath)
    {
        // Open Excel with no window prompts and create workbook instance
        Excel.Application app = new Excel.Application();
        app.DisplayAlerts = false;
        Excel.Workbook wb = app.Workbooks.Open(filepath);

        // Check if spreadsheet has information
        int count = wb.Worksheets.Count;
        if (count == 0)
        {
            Console.WriteLine("--> Spreadsheet has no sheets. Exclude spreadsheet from archiving");
        }

        // Remove data connections
        int count_conn = wb.Connections.Count;
        if (count_conn > 0)
        {
            for (int i = 1; i <= wb.Connections.Count; i++)
            {
                wb.Connections[i].Delete();
                i = i - 1;
            }
            count_conn = wb.Connections.Count;
            Console.WriteLine("--> Data connections detected and removed");
            wb.Save();
        }

        // Find and replace external cell chains with cell values
        bool hasChain = false;
        foreach (Excel.Worksheet sheet in wb.Sheets)
        {
            try
            {
                Excel.Range range = (Excel.Range)sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                foreach (Excel.Range cell in range.Cells)
                {
                    var value = cell.Value2;
                    string formula = cell.Formula.ToString();
                    string hit = formula.Substring(0, 2); // Transfer first 2 characters to string

                    if (hit == "='")
                    {
                        hasChain = true;
                        cell.Formula = "";
                        cell.Value2 = value;
                    }
                }
                if (hasChain == true)
                {
                    Console.WriteLine("--> External cell chains detected and replaced with cell values"); // Inform user
                    wb.Save(); // Save workbook
                }
            }
            catch (System.Runtime.InteropServices.COMException) // Catch if no formulas in range
            {
                // Do nothing
            }
            catch (System.ArgumentOutOfRangeException) // Catch if formula has less than 2 characters
            {
                // Do nothing
            }
        }

        // Find and replace RTD functions with cell values
        bool hasRTD = false;
        foreach (Excel.Worksheet sheet in wb.Sheets)
        {
            try
            {
                Excel.Range range = (Excel.Range)sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                foreach (Excel.Range cell in range.Cells)
                {
                    var value = cell.Value2;
                    string formula = cell.Formula.ToString();
                    string hit = formula.Substring(0, 4); // Transfer first 4 characters to string
                    if (hit == "=RTD")
                    {
                        cell.Formula = "";
                        cell.Value2 = value;
                        hasRTD = true;
                    }
                }
                if (hasRTD = true)
                {
                    Console.WriteLine("--> RTD functions detected and replaced with cell values");
                    wb.Save();
                }
            }
            catch (System.Runtime.InteropServices.COMException) // Catch if no formulas in range
            {
                // Do nothing
            }
        }
        
        try
        {
            // Make first sheet active
            if (app.Sheets.Count > 0)
            {
                Excel.Worksheet firstSheet = (Excel.Worksheet)app.ActiveWorkbook.Sheets[1];
                firstSheet.Activate();
                firstSheet.Select();
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Do nothing
        }

        // Save as .xlsx Strict
        new_filepath = Path.GetDirectoryName(filepath) + "\\1.xlsx"; //Rename with 1 and give extension .xlsx
        wb.SaveAs(new_filepath, 61); // 61 is code for Open XML Strict in Excel
        Console.WriteLine("--> Spreadsheet converted to .xlsx Strict conformance"); // Write user of conversion

        // Close Excel
        wb.Close();
        app.Quit();
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
        {
            Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
            Marshal.ReleaseComObject(app); // Delete Excel task in task manager
        }

        // Return success
        bool success = true;
        return success;
    }
}
