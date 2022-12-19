using System;
using System.IO;
using CommandLine;

namespace Convert.Spreadsheet
{
    class Program
    {
        public class Options
        {
            [Option('i', "inputfilepath", Required = true, HelpText = "The input filepath")]
            public string InputFilepath { get; set; }

            [Option('d', "delete", Required = false, HelpText = "Set to delete original file.")]
            public bool Delete { get; set; }

            [Option('r', "rename", Required = false, HelpText = "Define to rename output file.")]
            public string Rename { get; set; }

            [Option('p', "policy", Required = false, HelpText = "Set to convert file to comply with archival requirements.")]
            public bool Policy { get; set; }

            [Option('l', "libreoffice", Required = false, HelpText = "Set to use LibreOffice instead of Excel for conversion.")]

            public bool LibreOffice { get; set; }

            [Option('f', "outputfileformat", Required = true, HelpText = "Define output file format")]
            public string OutputFileFormat { get; set; }

            [Option('o', "outputfolder", Required = false, HelpText = "Define to save output file in custom folder. Default is same folder.")]
            public string OutputFolder { get; set; }
        }

        public static int Main(string[] args)
        {
            // Parse user arguments
            var parser = new Parser(with => with.HelpWriter = null);
            int exitcode = parser.ParseArguments<Options>(args).MapResult((opts) => RunApp(opts), errs => ShowHelp(errs));
            return exitcode;
        }

        static int RunApp(Options arg)
        {
            string input_extension = Path.GetExtension(arg.InputFilepath).ToLower();
            string output_extension = "." + arg.OutputFileFormat.ToLower().Split("-").First();
            string output_extension_LibreOffice = arg.OutputFileFormat.ToLower().Split("-").First();
            string output_folder, output_filepath;
            int fail = 0, success = 1;

            // Set output folder
            if (arg.OutputFolder != null && Directory.Exists(arg.OutputFolder))
            {
                output_folder = arg.OutputFolder;
            }
            else if (arg.OutputFolder != null && !Directory.Exists(arg.OutputFolder))
            {
                Console.WriteLine($"Output folder \"{arg.OutputFolder}\" does not exist");
                return fail;
            }
            else
            {
                output_folder = Path.GetDirectoryName(arg.InputFilepath);
            }

            // Set output filename         
            if (arg.Rename != null)
            {
                output_filepath = output_folder + "\\" + arg.Rename + output_extension;
            }
            else
            {
                output_filepath = output_folder + "\\" + Path.GetFileNameWithoutExtension(arg.InputFilepath) + output_extension;
            }

            // End program if no file exists
            if (!File.Exists(arg.InputFilepath))
            {
                Console.WriteLine("No file in input filepath");
                return fail;
            }

            // End program if input or output file formats are not accepted
            FileFormats check = new FileFormats();
            int input_check = check.CheckInputFileFormat(input_extension);
            if (input_check == 0)
            {
                Console.WriteLine("Input file format is not an accepted file format");
                return fail;
            }
            int output_check = check.CheckOutputFileFormat(output_extension);
            if (output_check == 0)
            {
                Console.WriteLine("Output file format is not an accepted file format");
                return fail;
            }

            // Define data types
            bool archive_success = false;
            bool convert_success = false;

            try
            {
                // Remove file attributes on file
                File.SetAttributes(arg.InputFilepath, FileAttributes.Normal);

                // Convert file
                Convert conversion = new Convert();
                // Use LibreOffice
                if (arg.LibreOffice == true) 
                {
                    convert_success = conversion.AnyFileFormat_LibreOffice(arg.InputFilepath, output_folder, output_extension_LibreOffice);
                }
                // Use LibreOffice, but file formats are not supported
                if (arg.LibreOffice == true && arg.OutputFileFormat == "xlsx-strict")
                {
                    Console.WriteLine("LibreOffice cannot convert to output file format");
                }
                // Use Excel
                else if (input_extension != ".numbers" || input_extension != ".fods") 
                {
                    // First transform file format to int
                    int xlFileFormat = check.ConvertFileFormatToInt(arg.OutputFileFormat);
                    // Then use int in conversion
                    convert_success = conversion.AnyFileFormat_Excel(arg.InputFilepath, output_filepath, xlFileFormat);
                }
                // Use Excel, but file formats are not supported
                else if (input_extension == ".numbers" || input_extension == ".fods") 
                {
                    Console.WriteLine("Excel cannot convert to output file format");
                }

                // Convert to comply with archival requirements
                if (arg.Policy == true)
                {
                    if (output_extension == ".xlsx")
                    {
                        // First repair file
                        Repair rep = new Repair();
                        rep.Repair_OOXML(output_filepath);

                        // Then comply with archival requirements
                        ArchiveRequirements ArcReq = new ArchiveRequirements();
                        archive_success = ArcReq.ArchiveRequirements_OOXML(output_filepath);
                    }
                    else
                    {
                        Console.WriteLine("File format policy compliance is only supported for XLSX output file format");
                    }
                }
            }

            // If spreadsheet is password protected or otherwise unreadable
            catch (FormatException)
            {
                Console.WriteLine("Input file cannot be read");
                return fail;
            }

            // Post conversion operations
            finally
            {
                if (convert_success == true)
                {
                    // If rename
                    if (arg.Rename != null && arg.LibreOffice == true)
                    {
                        string temp_filepath = output_folder + "\\" + Path.GetFileNameWithoutExtension(arg.InputFilepath) + output_extension;
                        File.Move(temp_filepath, output_filepath);
                    }

                    // If delete
                    if (arg.Delete == true)
                    {
                        // Delete original file, if filepath was not 1.xlsx
                        if (arg.InputFilepath != output_filepath)
                        {
                            File.Delete(arg.InputFilepath);
                            Console.WriteLine("Input file was deleted");
                        }
                    }
                }
            }
            // Return success to user
            Console.WriteLine("Program finished");
            return success;
        }

        // Show help to user, if parsing arguments fail
        static int ShowHelp(IEnumerable<Error> errs)
        {
            int fail = 0;
            Console.WriteLine("Input arguments have errors");
            return fail;
        }
    }
}
