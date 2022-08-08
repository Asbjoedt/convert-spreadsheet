using System;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    public string new_filepath = "";

    string Main(string[] args)
    {
        string filepath = args[0];
        string extension = Path.GetExtension(filepath);
        string message;

        try
        {
            switch (extension)
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
                    Console.WriteLine(filepath);

                    Convert_Requirements(filepath);

                    message = "File was successfully converted";

                    // Delete original file
                    if (filepath != new_filepath)
                    {
                        File.Delete(filepath);
                    }
                    return message;

                default:
                    message = "File format is not an accepted file format";
                    return message;
            }
        }
        catch (FileNotFoundException)
        {
            message = "No file in filepath";
            return message;
        }
        catch (FormatException)
        {
            message = "File is password-protected, editing-protected or corrupt";
            return message;
        }
    }

    void Convert_Requirements(string filepath)
    {
        // Open Excel
        Excel.Application app = new Excel.Application();
        app.DisplayAlerts = false;
        Excel.Workbook wb = app.Workbooks.Open(filepath);

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
        }

        // Save as .xlsx Strict
        new_filepath = Path.GetDirectoryName(filepath) + "\\1.xlsx";
        wb.SaveAs(new_filepath, 61); 
        Console.WriteLine("--> Spreadsheet converted to .xlsx Strict conformance");

        // Close Excel
        wb.Close();
        app.Quit();
    }
}
