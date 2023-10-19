using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Office2013.ExcelAc;
using ImageMagick;
using DocumentFormat.OpenXml;

namespace Convert.Spreadsheet
{
    public class Policy
    {
        public void OOXML_Errors(string filepath)
        {
            Remove_DataConnections(filepath);
            Remove_CellReferences(filepath);
            Remove_RTDFunctions(filepath);
            Remove_ExternalObjects(filepath);
            Convert_EmbeddedImages(filepath);
        }
        public void OOXML_Warnings(string filepath)
        {
            Change_Conformance_ExcelInterop(filepath);
            Remove_AbsolutePath(filepath);
            Activate_FirstSheet(filepath);
        }

        // Change conformance to Strict
        public void Change_Conformance_ExcelInterop(string filepath)
        {
            // Open Excel
            Excel.Application app = new Excel.Application(); // Create Excel object instance
            app.DisplayAlerts = false; // Don't display any Excel prompts
            Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

            // Convert to Strict and close Excel
            wb.SaveAs(filepath, 61);
            wb.Close();
            app.Quit();

            // If CLISC is run on Windows close Excel in task manager
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Marshal.ReleaseComObject(wb); // Delete workbook task
                Marshal.ReleaseComObject(app); // Delete Excel task
            }
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

                // Delete all QueryTableParts
                IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    // Delete all QueryTableParts in WorksheetParts
                    List<QueryTablePart> queryTables = worksheetPart.QueryTableParts.ToList(); // Must be a list
                    foreach (QueryTablePart queryTablePart in queryTables)
                    {
                        worksheetPart.DeletePart(queryTablePart);
                    }

                    // Delete all QueryTableParts, if they are not registered in a WorksheetPart
                    List<TableDefinitionPart> tableDefinitionParts = worksheetPart.TableDefinitionParts.ToList();
                    foreach (TableDefinitionPart tableDefinitionPart in tableDefinitionParts)
                    {
                        List<IdPartPair> idPartPairs = tableDefinitionPart.Parts.ToList();
                        foreach (IdPartPair idPartPair in idPartPairs)
                        {
                            if (idPartPair.OpenXmlPart.ToString() == "DocumentFormat.OpenXml.Packaging.QueryTablePart")
                            {
                                // Delete QueryTablePart
                                tableDefinitionPart.DeletePart(idPartPair.OpenXmlPart);
                                // The TableDefinitionPart must also be deleted
                                worksheetPart.DeletePart(tableDefinitionPart);
                                // And the reference to the TableDefinitionPart in the WorksheetPart must be deleted
                                List<TablePart> tableParts = worksheetPart.Worksheet.Descendants<TablePart>().ToList();
                                foreach (TablePart tablePart in tableParts)
                                {
                                    if (idPartPair.RelationshipId == tablePart.Id)
                                    {
                                        tablePart.Remove();
                                    }
                                }
                            }
                        }
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
                IEnumerable<ExternalWorkbookPart> extWbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts;
                foreach (ExternalWorkbookPart extWbPart in extWbParts)
                {
                    List<ExternalRelationship> extrels = extWbPart.ExternalRelationships.ToList(); // Must be a list
                    foreach (ExternalRelationship extrel in extrels)
                    {
                        Uri uri = new Uri($"External reference {extrel.Uri} was removed", UriKind.Relative);
                        extWbPart.DeleteExternalRelationship(extrel.Id);
                        extWbPart.AddExternalRelationship(relationshipType: "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject", externalUri: uri, id: extrel.Id);
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
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true, new OpenSettings()
            {
                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013)
            }))
            {
                if (spreadsheet.WorkbookPart.Workbook.AbsolutePath != null)
                {
                    AbsolutePath absPath = spreadsheet.WorkbookPart.Workbook.AbsolutePath;
                    absPath.Remove();
                    Console.WriteLine("Absolute path to local directory removed");
                }
            }
        }

        public void Convert_EmbeddedImages(string filepath)
        {
            List<ImagePart> emf = new List<ImagePart>();
            List<ImagePart> images = new List<ImagePart>();

            // Open spreadsheet
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    // Perform check
                    emf = worksheetPart.ImageParts.Distinct().ToList();
                    if (worksheetPart.DrawingsPart != null) // DrawingsPart needs a null check
                    {
                        images = worksheetPart.DrawingsPart.ImageParts.Distinct().ToList();
                    }

                    // Perform change

                    // Convert Excel-generated .emf images to TIFF
                    foreach (ImagePart imagePart in emf)
                    {
                        Convert_EmbedEmf(filepath, worksheetPart, imagePart);
                    }

                    // Convert embedded images to TIFF
                    foreach (ImagePart imagePart in images)
                    {
                        Convert_EmbedImg(filepath, worksheetPart, imagePart);
                    }
                }
            }
        }

        // Convert embedded images to TIFF
        public void Convert_EmbedImg(string filepath, WorksheetPart worksheetPart, ImagePart imagePart)
        {
            // Convert streamed image to new stream
            Stream stream = imagePart.GetStream();
            Stream new_Stream = Convert_ImageMagick(stream);
            stream.Dispose();

            // Add new ImagePart
            ImagePart new_ImagePart = worksheetPart.DrawingsPart.AddImagePart(ImagePartType.Tiff);

            // Save image from stream to new ImagePart
            new_Stream.Position = 0;
            new_ImagePart.FeedData(new_Stream);

            // Change relationships of image
            string id = Get_RelationshipId(imagePart);
            Blip blip = worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>()
                            .Where(p => p.BlipFill.Blip.Embed == id)
                            .Select(p => p.BlipFill.Blip)
                            .Single();
            blip.Embed = Get_RelationshipId(new_ImagePart);

            // Delete original ImagePart
            worksheetPart.DrawingsPart.DeletePart(imagePart);
        }

        // Convert Excel-generated .emf images to TIFF
        public void Convert_EmbedEmf(string filepath, WorksheetPart worksheetPart, ImagePart imagePart)
        {
            // Convert streamed image to new stream
            Stream stream = imagePart.GetStream();
            Stream new_Stream = Convert_ImageMagick(stream);
            stream.Dispose();

            // Add new ImagePart
            ImagePart new_ImagePart = worksheetPart.VmlDrawingParts.First().AddImagePart(ImagePartType.Tiff);

            // Save image from stream to new ImagePart
            new_Stream.Position = 0;
            new_ImagePart.FeedData(new_Stream);

            // Change relationships of image
            string id = Get_RelationshipId(imagePart);
            XDocument xElement = worksheetPart.VmlDrawingParts.First().GetXDocument();
            IEnumerable<XElement> descendants = xElement.FirstNode.Document.Descendants();
            foreach (XElement descendant in descendants)
            {
                if (descendant.Name == "{urn:schemas-microsoft-com:vml}imagedata")
                {
                    IEnumerable<XAttribute> attributes = descendant.Attributes();
                    foreach (XAttribute attribute in attributes)
                    {
                        if (attribute.Name == "{urn:schemas-microsoft-com:office:office}relid")
                        {
                            if (attribute.Value == id)
                            {
                                attribute.Value = Get_RelationshipId(new_ImagePart);
                                worksheetPart.VmlDrawingParts.First().SaveXDocument();
                            }
                        }
                    }
                }
            }
            // Delete original ImagePart
            worksheetPart.VmlDrawingParts.First().DeletePart(imagePart);
        }

        // Convert embedded object to TIFF using ImageMagick
        public Stream Convert_ImageMagick(Stream stream)
        {
            // Read the input stream in ImageMagick
            using (MagickImage image = new MagickImage(stream))
            {
                // Set input stream position to beginning
                stream.Position = 0;

                // Create a memorystream to write image to
                MemoryStream new_stream = new MemoryStream();

                // Adjust TIFF settings
                image.Format = MagickFormat.Tiff;
                image.Settings.ColorSpace = ColorSpace.RGB;
                image.Settings.Depth = 32;
                image.Settings.Compression = CompressionMethod.LZW;

                // Write image to stream
                image.Write(new_stream);

                // Return the memorystream
                return new_stream;
            }
        }

        // Get relationship id of an OpenXmlPart
        public string Get_RelationshipId(OpenXmlPart part)
        {
            string id = "";
            IEnumerable<OpenXmlPart> parentParts = part.GetParentParts();
            foreach (OpenXmlPart parentPart in parentParts)
            {
                if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.DrawingsPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.VmlDrawingPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.Model3DReferenceRelationshipPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.OleObjectPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
            }
            return id;
        }

        // If file is ODS, use external app
        public void ODS(string filepath, bool strict)
        {
            Process app = new Process();
            app.StartInfo.UseShellExecute = false;
            app.StartInfo.FileName = "javaw";

            // If app is run on Windows
            string? dir = null;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                dir = Environment.GetEnvironmentVariable("ODS-ArchivalRequirements");
            }
            if (dir != null)
            {
                app.StartInfo.FileName = dir;
            }
            else
            {
                app.StartInfo.FileName = "C:\\Program Files\\ODS-ArchivalRequirements\\ODS-ArchivalRequirements.jar";
            }
            if (strict)
            {
                app.StartInfo.Arguments = $"--inputfilepath \"{filepath}\" --change --policy-strict";
            }
            else
            {
                app.StartInfo.Arguments = $"--inputfilepath \"{filepath}\" --change";
            }            
            app.Start();
            app.WaitForExit();
            app.Close();
        }
    }
}
