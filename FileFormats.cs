using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Convert.Spreadsheet
{
    public class FileFormats
    {
        public int CheckInputFileFormat(string input_extension)
        {
            int fail = 0, success = 1;

            string[] input_fileformats = { ".fods", ".numbers", ".ods", ".ots", "xla", ".xls", ".xlt", "xlam", ".xlsb", ".xlsm", ".xlsx", ".xltm", ".xltx" };
            if (!input_fileformats.Contains(input_extension))
            {
                return fail;
            }
            return success;
        }

        public int CheckOutputFileFormat(string output_fileformat)
        {
            int fail = 0, success = 1;

            string[] output_fileformats = { ".fods", ".ods", ".ots", ".xlsb", ".xlsm", ".xlsx", ".xltm", ".xltx", ".csv", ".html", ".mht", ".txt" };
            if (!output_fileformats.Contains(output_fileformat))
            {
                return fail;
            }
            return success;
        }

        public int ConvertFileFormatToInt(string output_fileformat)
        {
            int output = 0;

            switch (output_fileformat)
            {
                case "ods":
                    return 60;
                case "xlsx-strict":
                    return 61;
                case "xlsx":
                    return 51;
                case "xlam":
                    return 55;
                case "xltx":
                    return 54;
                case "xltm":
                    return 53;
                case "xlsm":
                    return 52;
                case "xlsb":
                    return 50;
                case "xla":
                    return 18;
                case "xls":
                    return 56;
                case "xlt":
                    return 17;
                case "csv":
                    return 6;
                case "html":
                    return 44;
                case "mht":
                    return 45;
                case "txt":
                    return 42;
            }
            return output;
        }
    }
}