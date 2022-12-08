using System;
using System.IO;
using CommandLine;

namespace convert_spreadsheet
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

            [Option('o', "outputfolder", Required = false, HelpText = "Define to save output file in custom folder. Default is same folder.")]
            public string OutputFolder { get; set; }
        }

        public static void Main(string[] args)
        {
            // Parse user arguments
            var parser = new Parser(with => with.HelpWriter = null);
            var parse_args = parser.ParseArguments<Options>(args);
            parse_args
            .WithParsed(RunApp);
        }

        public static void RunApp(Options arg)
        {
            string input_extension = Path.GetExtension(arg.InputFilepath).ToLower();
            string output_folder;
            string output_filepath;

            // Write filepath to user
            Console.WriteLine($"Input filepath: {arg.InputFilepath}");

            // Set output folder
            if (arg.OutputFolder != null && Directory.Exists(arg.OutputFolder))
            {
                output_folder = arg.OutputFolder;
            }
            else if (arg.OutputFolder != null && !Directory.Exists(arg.OutputFolder))
            {
                Console.WriteLine($"Output folder \"{arg.OutputFolder}\" does not exist");
                throw new DirectoryNotFoundException("Output folder does not exist");
            }
            else
            {
                output_folder = Path.GetDirectoryName(arg.InputFilepath);
            }

            // Set output filename
            if (arg.Rename != null)
            {
                output_filepath = output_folder + "\\" + arg.Rename + ".xlsx";
            }
            else
            {
                output_filepath = output_folder + "\\" + Path.GetFileNameWithoutExtension(arg.InputFilepath) + ".xlsx";
            }

            Convert conversion = new Convert();
            ArchiveRequirements ArcReq = new ArchiveRequirements();
            bool convert_success = false;
            bool archive_success = false;

            // Quit program if no file exists
            if (!File.Exists(arg.InputFilepath))
            {
                Console.WriteLine("No file in input filepath");
                Environment.Exit(0);
            }

            try
            {
                // The switch includes all accepted file extensions for conversion
                switch (input_extension)
                {
                    case ".ods":
                    case ".ots":
                    case ".xls":
                    case ".xlt":
                    case ".xlsb":
                        // Convert spreadsheet to .xlsx
                        convert_success = conversion.Convert_All(arg.InputFilepath, output_filepath);

                        // Comply with archiving requirements
                        archive_success = ArcReq.ArchiveRequirements_OOXML(output_filepath);
                        break;

                    case ".xlsm":
                    case ".xlsx":
                    case ".xltm":
                    case ".xltx":
                        // Convert spreadsheet to .xlsx
                        convert_success = conversion.Convert_OOXML(arg.InputFilepath, output_filepath);

                        // Comply with archiving requirements
                        archive_success = ArcReq.ArchiveRequirements_OOXML(output_filepath);
                        break;

                    default:
                        // If the filepath has extension not included in switch
                        Console.WriteLine("File format is not an accepted file format");
                        break;
                }
            }

            // If spreadsheet is password protected or otherwise unreadable
            catch (FormatException)
            {
                Console.WriteLine("Input file cannot be read");
            }

            // Post conversion operations
            finally
            {
                if (convert_success == true && archive_success == true)
                {
                    if (arg.Delete == true)
                    {
                        // Delete original file, if filepath was not 1.xlsx
                        if (arg.InputFilepath != output_filepath)
                        {
                            File.Delete(arg.InputFilepath);
                            Console.WriteLine("Input file was deleted");
                        }
                    }

                    // Write output filepath to user
                    Console.WriteLine("Output filepath is: " + output_filepath);
                }
            }
        }
    }
}
