using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2013.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace convert_spreadsheet
{
    public class ArchiveRequirements
    {
        public bool ArchiveRequirements_OOXML(string filepath)
        {
            bool success = false;

            Remove_DataConnections(filepath);
            Remove_CellReferences(filepath);
            Remove_RTDFunctions(filepath);
            Remove_PrinterSettings(filepath);
            Remove_ExternalObjects(filepath);
            Activate_FirstSheet(filepath);
            Remove_AbsolutePath(filepath);

            // Inform user and return success
            Console.WriteLine("File complies with archival requirements");
            success = true;
            return success;
        }

        // Remove data connections
        public void Remove_DataConnections(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                // Delete all connections
                ConnectionsPart conn = spreadsheet.WorkbookPart.ConnectionsPart;
                if (conn != null)
                {
                    spreadsheet.WorkbookPart.DeletePart(conn);
                    Console.WriteLine("Data connection was removed");
                }

                // Delete all query tables
                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    List<QueryTablePart> queryTables = part.QueryTableParts.ToList();
                    foreach (QueryTablePart qtp in queryTables)
                    {
                        part.DeletePart(qtp);
                    }
                }

                // If spreadsheet contains a custom XML Map, delete databinding
                if (spreadsheet.WorkbookPart.CustomXmlMappingsPart != null)
                {
                    CustomXmlMappingsPart xmlMap = spreadsheet.WorkbookPart.CustomXmlMappingsPart;
                    List<Map> maps = xmlMap.MapInfo.Elements<Map>().ToList();
                    foreach (Map map in maps)
                    {
                        if (map.DataBinding != null)
                        {
                            map.DataBinding.Remove();
                        }
                    }
                }
            }
            // Repair spreadsheet
            Repair rep = new Repair();
            //rep.Repair_QueryTables(filepath);
        }

        // Remove RTD functions
        public void Remove_RTDFunctions(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    Worksheet worksheet = part.Worksheet;
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                    foreach (var row in rows)
                    {
                        var cells = row.Elements<Cell>();
                        foreach (Cell cell in cells)
                        {
                            if (cell.CellFormula != null)
                            {
                                string formula = cell.CellFormula.InnerText;
                                if (formula.Length > 2)
                                {
                                    string hit = formula.Substring(0, 3); // Transfer first 3 characters to string
                                    if (hit == "RTD")
                                    {
                                        CellValue cellvalue = cell.CellValue; // Save current cell value
                                        cell.CellFormula = null; // Remove RTD formula
                                        // If cellvalue does not have a real value
                                        if (cellvalue.Text == "#N/A")
                                        {
                                            cell.DataType = CellValues.String;
                                            cell.CellValue = new CellValue("Invalid data removed");
                                        }
                                        else
                                        {
                                            cell.CellValue = cellvalue; // Insert saved cell value
                                        }
                                        Console.WriteLine("RTD function was removed");
                                    }
                                }
                            }
                        }
                    }
                }
                // Delete calculation chain
                CalculationChainPart calc = spreadsheet.WorkbookPart.CalculationChainPart;
                spreadsheet.WorkbookPart.DeletePart(calc);

                // Delete volatile dependencies
                VolatileDependenciesPart vol = spreadsheet.WorkbookPart.VolatileDependenciesPart;
                spreadsheet.WorkbookPart.DeletePart(vol);
            }
        }

        // Remove printer settings
        public void Remove_PrinterSettings(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                List<WorksheetPart> wsParts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart wsPart in wsParts)
                {
                    List<SpreadsheetPrinterSettingsPart> printerList = wsPart.SpreadsheetPrinterSettingsParts.ToList();
                    foreach (SpreadsheetPrinterSettingsPart printer in printerList)
                    {
                        wsPart.DeletePart(printer);
                        Console.WriteLine("Printer setting was removed");
                    }
                }
            }
        }

        // Remove external cell references
        public void Remove_CellReferences(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                foreach (WorksheetPart part in worksheetparts)
                {
                    Worksheet worksheet = part.Worksheet;
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                    foreach (var row in rows)
                    {
                        var cells = row.Elements<Cell>();
                        foreach (Cell cell in cells)
                        {
                            if (cell.CellFormula != null)
                            {
                                string formula = cell.CellFormula.InnerText;
                                if (formula.Length > 1)
                                {
                                    string hit = formula.Substring(0, 1); // Transfer first 1 characters to string
                                    string hit2 = formula.Substring(0, 2); // Transfer first 2 characters to string
                                    if (hit == "[" || hit2 == "'[")
                                    {
                                        CellValue cellvalue = cell.CellValue; // Save current cell value
                                        cell.CellFormula = null;
                                        // If cellvalue does not have a real value
                                        if (cellvalue.Text == "#N/A")
                                        {
                                            cell.DataType = CellValues.String;
                                            cell.CellValue = new CellValue("Invalid data removed");
                                        }
                                        else
                                        {
                                            cell.CellValue = cellvalue; // Insert saved cell value
                                        }
                                        Console.WriteLine("External cell reference was removed");
                                    }
                                }
                            }
                        }
                    }
                }

                // Delete external book references
                List<ExternalWorkbookPart> extwbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts.ToList();
                if (extwbParts.Count > 0)
                {
                    foreach (ExternalWorkbookPart extpart in extwbParts)
                    {
                        var elements = extpart.ExternalLink.ChildElements.ToList();
                        foreach (var element in elements)
                        {
                            if (element.LocalName == "externalBook")
                            {
                                spreadsheet.WorkbookPart.DeletePart(extpart);
                            }
                        }
                    }
                }

                // Delete calculation chain
                CalculationChainPart calc = spreadsheet.WorkbookPart.CalculationChainPart;
                spreadsheet.WorkbookPart.DeletePart(calc);

                // Delete defined names that includes external cell references
                DefinedNames definedNames = spreadsheet.WorkbookPart.Workbook.DefinedNames;
                if (definedNames != null)
                {
                    var definedNamesList = definedNames.ToList();
                    foreach (DefinedName definedName in definedNamesList)
                    {
                        if (definedName.InnerXml.StartsWith("["))
                        {
                            definedName.Remove();
                        }
                    }
                }
            }
        }

        // Remove external object references
        public void Remove_ExternalObjects(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                List<ExternalWorkbookPart> extwbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts.ToList();
                if (extwbParts.Count > 0)
                {
                    foreach (ExternalWorkbookPart extpart in extwbParts)
                    {
                        if (extpart.ExternalLink.ChildElements != null)
                        {
                            var elements = extpart.ExternalLink.ChildElements.ToList();
                            foreach (var element in elements)
                            {
                                if (element.LocalName == "oleLink")
                                {
                                    spreadsheet.WorkbookPart.DeletePart(extpart);
                                    Console.WriteLine("External object reference was removed");
                                }
                            }
                        }
                    }
                }
            }
        }

        // Make first sheet active sheet
        public void Activate_FirstSheet(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                BookViews bookViews = spreadsheet.WorkbookPart.Workbook.GetFirstChild<BookViews>();
                WorkbookView workbookView = bookViews.GetFirstChild<WorkbookView>();
                if (workbookView.ActiveTab != null)
                {
                    var activeSheetId = workbookView.ActiveTab.Value;
                    if (activeSheetId > 0)
                    {
                        workbookView.ActiveTab.Value = 0;

                        List<WorksheetPart> worksheets = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                        foreach (WorksheetPart worksheet in worksheets)
                        {
                            var sheetviews = worksheet.Worksheet.SheetViews.ToList();
                            foreach (SheetView sheetview in sheetviews)
                            {
                                sheetview.TabSelected = null;
                                Console.WriteLine("First sheet was activated");
                            }
                        }
                    }
                }
            }
        }

        // Remove absolute path to local directory
        public void Remove_AbsolutePath(string filepath)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                if (spreadsheet.WorkbookPart.Workbook.AbsolutePath != null)
                {
                    AbsolutePath absPath = spreadsheet.WorkbookPart.Workbook.GetFirstChild<AbsolutePath>();
                    absPath.Remove();
                    Console.WriteLine("Absolute path to local directory removed");
                }
            }
        }
    }
}
