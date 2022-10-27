using System;
using System.IO;

namespace convert_spreadsheet
{
    class Program
    {
        public static void Main(string[] args)
        {
            string input_filepath = args[0];
            string input_extension = Path.GetExtension(input_filepath);
            string output_filepath = Path.GetDirectoryName(input_filepath) + "\"1.xlsx";

            try
            {
                switch (input_extension) // The switch includes all accepted file extensions
                {
                    case ".ods":
                    case ".ODS":
                    case ".ots":
                    case ".OTS":
                    case ".xls":
                    case ".XLS":
                    case ".xlsb":
                    case ".XLSB":
                    case ".xlsm":
                    case ".XLSM":
                    case ".xlsx":
                    case ".XLSX":
                    case ".xlt":
                    case ".XLT":
                    case ".xltm":
                    case ".XLTM":
                    case ".xltx":
                    case ".XLTX":
                        // Write filepath to user
                        Console.WriteLine(input_filepath);

                        // Convert spreadsheet to .xlsx
                        Convert conversion = new Convert();
                        bool convert_success = conversion.Convert_All(input_filepath, output_filepath);

                        // Write output filepath to user
                        Console.WriteLine("New filepath is: " + output_filepath);

                        // Comply with archiving requirements
                        ArchiveRequirements ArcReq = new ArchiveRequirements();
                        bool archive_success = ArcReq.ArchiveRequirements_OOXML(output_filepath);

                        // Delete original file, if filepath was not 1.xlsx
                        if (input_filepath != output_filepath) 
                        {
                            File.Delete(input_filepath);
                            Console.WriteLine("Input file was deleted");
                        }
                        break;

                    default:
                        // If the filepath has extension not included in switch
                        Console.WriteLine("File format is not an accepted file format");
                        break;
                }
            }
            // If filepath has not file
            catch (FileNotFoundException) 
            {
                Console.WriteLine("No file in filepath");
            }
            // If spreadsheet is password protected or otherwise unreadable
            catch (FormatException) 
            {
                File.Delete(input_filepath); // Delete file
                Console.WriteLine("File cannot be read");
            }
        }
    }
}
