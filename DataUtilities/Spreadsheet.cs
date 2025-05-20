using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using System.Text.RegularExpressions;
using PatternFill = DocumentFormat.OpenXml.Spreadsheet.PatternFill;
using Fill = DocumentFormat.OpenXml.Spreadsheet.Fill;
using System.ComponentModel.DataAnnotations;
using DbContextClassLibrary.Models;

namespace DataUtilities
{
    public class Spreadsheet
    {
        /// <summary>  
        /// Extracts Excel data from a file path.  
        /// </summary>  
        public static (List<List<SheetRow>>, List<string>) ExtractExcelData(string filePath)
        {
            var returnSheets = new List<List<SheetRow>>();
            List<string> sheetTitles = new List<string>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                ProcessSheet(returnSheets, sheetTitles, document);
            }

            return (returnSheets, sheetTitles);
        }

        public class SheetRow
        {
            public int Row { get; set; }
            public List<Dictionary<string, object>> Cells { get; set; }
        }

        /// <summary>  
        /// Extracts Excel data from a file stream.  
        /// </summary>  
        public static (List<List<SheetRow>>, List<string>) ExtractExcelData(Stream fileStream)
        {
            var returnSheets = new List<List<SheetRow>>();
            List<string> sheetTitles = new List<string>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileStream, false))
            {
                ProcessSheet(returnSheets, sheetTitles, document);
            }

            return (returnSheets, sheetTitles);
        }

        /// <summary>  
        /// Processes each sheet in the Excel document.  
        /// </summary>  
        private static void ProcessSheet(List<List<SheetRow>> returnSheets, List<string> sheetTitles, SpreadsheetDocument document)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            var sheets = workbookPart.Workbook.Sheets;

            foreach (Sheet sheet in sheets)
            {
                // Exclude hidden sheets  
                if (sheet.State != null && sheet.State.Value == SheetStateValues.Hidden)
                {
                    Console.WriteLine($"Skipping hidden sheet: {sheet.Name}");
                    continue;
                }

                var result = new List<SheetRow>();

                // Get worksheet part  
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Handle merged cells  
                var mergeCells = worksheetPart.Worksheet.Descendants<MergeCell>();
                var mergedCellMap = CreateMergedCellMap(worksheetPart, workbookPart, mergeCells);

                // Determine the maximum number of columns  
                int maxColumns = 0;
                foreach (var row in sheetData.Elements<Row>())
                {
                    foreach (var cell in row.Elements<Cell>())
                    {
                        int columnIndex = GetColumnIndexFromCellReference(cell.CellReference.Value);
                        if (columnIndex > maxColumns)
                            maxColumns = columnIndex;
                    }
                }

                // Read each row  
                int rowNumber = 0;
                foreach (var row in sheetData.Elements<Row>())
                {
                    // Initialize rowData with empty dictionaries  
                    var rowData = new SheetRow()
                    {
                        Row = rowNumber,
                        Cells = new List<Dictionary<string, object>>()
                    };
                    for (int i = 0; i < maxColumns; i++)
                    {
                        rowData.Cells.Add(new Dictionary<string, object>
                        {
                            { "Value", "" },
                            { "Col", i }
                        });
                    }

                    foreach (var cell in row.Elements<Cell>())
                    {
                        int columnIndex = GetColumnIndexFromCellReference(cell.CellReference.Value) - 1; // Zero-based index  

                        string cellReference = cell.CellReference.Value;

                        string cellValue = GetCellValue(cell, workbookPart);

                        // Check if this cell is part of a merged range  
                        if (mergedCellMap.TryGetValue(cellReference, out string mergedValue))
                        {
                            cellValue = mergedValue;
                        }

                        // Assign the value to the correct position in rowData  
                        if (columnIndex >= 0 && columnIndex < rowData.Cells.Count)
                        {
                            rowData.Cells[columnIndex]["Value"] = string.IsNullOrEmpty(cellValue) ? "" : cellValue;

                            // Extract formatting information  
                            var formatting = GetCellFormatting(cell, workbookPart);
                            if (formatting.Count > 0)
                            {
                                rowData.Cells[columnIndex]["Formatting"] = formatting;
                            }
                        }
                        else
                        {
                            // Optionally handle out-of-range columns  
                            // For now, we'll skip them  
                        }
                    }

                    // After processing all cells in the row, handle merged cells that might not have explicit cell entries  
                    foreach (var kvp in mergedCellMap)
                    {
                        string cellRef = kvp.Key;
                        string value = kvp.Value;

                        var match = Regex.Match(cellRef, @"^([A-Z]+)(\d+)$");
                        if (match.Success)
                        {
                            string columnLetters = match.Groups[1].Value;
                            int rowNum = int.Parse(match.Groups[2].Value);

                            if (rowNum == row.RowIndex)
                            {
                                int colIndex = GetColumnIndexFromColumnName(columnLetters) - 1; // Zero-based  
                                if (colIndex >= 0 && colIndex < rowData.Cells.Count)
                                {
                                    if (string.IsNullOrEmpty(rowData.Cells[colIndex]["Value"] as string))
                                    {
                                        rowData.Cells[colIndex]["Value"] = value;
                                    }
                                }
                            }
                        }
                    }

                    result.Add(rowData);
                    rowNumber++;
                }
                returnSheets.Add(result);
                sheetTitles.Add(sheet.Name);
            }
        }

        /// <summary>  
        /// Retrieves the cell formatting information as a dictionary containing only the present formatting attributes, with colors in hex format.  
        /// </summary>  
        private static Dictionary<string, object> GetCellFormatting(Cell cell, WorkbookPart workbookPart)
        {
            var formatting = new Dictionary<string, object>();

            if (cell.StyleIndex != null)
            {
                var styles = workbookPart.WorkbookStylesPart?.Stylesheet;
                if (styles != null)
                {
                    // Ensure StyleIndex is within the range of CellFormats  
                    if ((int)cell.StyleIndex.Value < styles.CellFormats.Count())
                    {
                        // Cast the CellFormat to access specific properties  
                        CellFormat cellFormat = styles.CellFormats.ElementAt((int)cell.StyleIndex.Value) as CellFormat;

                        if (cellFormat != null)
                        {
                            // Check borders  
                            if (cellFormat.BorderId != null && (int)cellFormat.BorderId.Value < styles.Borders.Count())
                            {
                                Border border = styles.Borders.ElementAt((int)cellFormat.BorderId.Value) as Border;
                                if (border != null && (
                                    border.LeftBorder != null ||
                                    border.RightBorder != null ||
                                    border.TopBorder != null ||
                                    border.BottomBorder != null ||
                                    border.DiagonalBorder != null))
                                {
                                    formatting["Border"] = true;
                                }
                            }

                            // Check fill (background color)  
                            if (cellFormat.FillId != null && (int)cellFormat.FillId.Value < styles.Fills.Count())
                            {
                                Fill fill = styles.Fills.ElementAt((int)cellFormat.FillId.Value) as Fill;
                                if (fill != null)
                                {
                                    PatternFill patternFill = fill.PatternFill;
                                    if (patternFill != null && patternFill.ForegroundColor != null)
                                    {
                                        string hexColor = "None";

                                        // Priority: RGB > Theme > Indexed  
                                        if (patternFill.ForegroundColor.Rgb != null)
                                        {
                                            // Use RGB value directly (assumes ARGB; trim to RGB)  
                                            string rgb = patternFill.ForegroundColor.Rgb.Value;
                                            if (rgb.Length == 8) // ARGB  
                                                rgb = rgb.Substring(2, 6); // Extract RGB  
                                            hexColor = $"#{rgb.ToUpper()}";
                                        }
                                        else if (patternFill.ForegroundColor.Theme != null && workbookPart.ThemePart != null)
                                        {
                                            // Resolve Theme color to hex  
                                            hexColor = GetThemeColorHex((int)patternFill.ForegroundColor.Theme.Value, workbookPart.ThemePart);
                                        }
                                        else if (patternFill.ForegroundColor.Indexed != null)
                                        {
                                            // Map Indexed color to hex  
                                            hexColor = GetIndexedColorHex((int)patternFill.ForegroundColor.Indexed.Value);
                                        }

                                        if (!string.IsNullOrEmpty(hexColor) && hexColor != "None")
                                        {
                                            formatting["Color"] = hexColor;
                                        }
                                    }
                                }
                            }

                            // Check font (bold)  
                            if (cellFormat.FontId != null && (int)cellFormat.FontId.Value < styles.Fonts.Count())
                            {
                                Font font = styles.Fonts.ElementAt((int)cellFormat.FontId.Value) as Font;
                                if (font != null)
                                {
                                    Bold bold = font.Bold;
                                    if (bold != null && (bold.Val == null || bold.Val.Value)) // If Val is null, Bold is considered true  
                                    {
                                        formatting["Bold"] = true;
                                    }

                                    // Optionally, check for italic, underline, etc.  
                                }
                            }
                        }
                    }
                }
            }

            return formatting;
        }

        /// <summary>  
        /// Maps a Theme color index to its corresponding hex color code using the workbook's theme.  
        /// </summary>  
        private static string GetThemeColorHex(int themeColorIndex, ThemePart themePart)
        {
            if (themePart.Theme == null)
            {
                Console.WriteLine("ThemePart does not contain a Theme.");
                return "None";
            }

            // Access the ThemeElements directly  
            var themeElements = themePart.Theme.ThemeElements;
            if (themeElements == null)
            {
                Console.WriteLine("Theme does not contain ThemeElements.");
                return "None";
            }

            // Access the ColorScheme within ThemeElements  
            var colorScheme = themeElements.ColorScheme;
            if (colorScheme == null)
            {
                Console.WriteLine("ThemeElements do not contain a ColorScheme.");
                return "None";
            }

            // Retrieve all SchemeColor elements  
            var schemeColors = colorScheme.Elements<SchemeColor>().ToList();
            if (themeColorIndex < 0 || themeColorIndex >= schemeColors.Count)
            {
                Console.WriteLine($"ThemeColorIndex {themeColorIndex} is out of range. Total scheme colors: {schemeColors.Count}");
                return "None";
            }

            var schemeColor = schemeColors[themeColorIndex];

            // Attempt to retrieve the RGB color  
            var rgbColor = schemeColor.GetFirstChild<RgbColorModelHex>();
            if (rgbColor != null && !string.IsNullOrEmpty(rgbColor.Val))
            {
                string rgb = rgbColor.Val;
                if (rgb.Length == 8) // ARGB  
                    rgb = rgb.Substring(2, 6); // Extract RGB  
                return $"#{rgb.ToUpper()}";
            }

            // If RGB is not available, attempt to handle SystemColor or Theme-based colors as needed  
            var sysColor = schemeColor.GetFirstChild<SystemColor>();
            if (sysColor != null)
            {
                // Handle system colors if necessary  
                // For simplicity, return "None"  
                Console.WriteLine($"SystemColor found for themeColorIndex {themeColorIndex}, but system colors are not handled.");
                return "None";
            }

            // If no color information is found, return "None"  
            Console.WriteLine($"No RGB or SystemColor found for themeColorIndex {themeColorIndex}.");
            return "None";
        }

        /// <summary>  
        /// Maps an Indexed color index to its corresponding hex color code using Excel's default palette.  
        /// </summary>  
        private static string GetIndexedColorHex(int indexedColor)
        {
            // Excel's default Indexed color palette (0-56)  
            // This is a subset; extend as needed  
            string[] indexedColors = new string[]
            {
                "#000000", // 0 Black  
                "#FFFFFF", // 1 White  
                "#FF0000", // 2 Red  
                "#00FF00", // 3 Lime  
                "#0000FF", // 4 Blue  
                "#FFFF00", // 5 Yellow  
                "#FF00FF", // 6 Magenta  
                "#00FFFF", // 7 Cyan  
                "#800000", // 8 Maroon  
                "#808000", // 9 Olive  
                "#008000", // 10 Green  
                "#800080", // 11 Purple  
                "#008080", // 12 Teal  
                "#000080", // 13 Navy  
                "#FFA500", // 14 Orange  
                "#808080", // 15 Gray  
                "#C0C0C0", // 16 Silver  
                "#9999FF", // 17 Light Blue  
                "#993366", // 18 Midnight Blue  
                "#FFFFCC", // 19 Very Light Yellow  
                "#CCFFFF", // 20 Pale Aqua  
                "#660066", // 21 Plum  
                "#FF8080", // 22 Light Red  
                "#66FF66", // 23 Light Green  
                "#8080FF", // 24 Light Blue  
                "#9966FF", // 25 Amethyst  
                "#FFB266", // 26 Peach  
                "#B266FF", // 27 Lavender  
                "#4DA6FF", // 28 Sky Blue  
                "#808000", // 29 Olive  
                "#FF4D4D", // 30 Coral  
                "#99E6E6", // 31 Aqua  
                "#6666FF", // 32 Royal Blue  
                "#FF99E6", // 33 Pink  
                "#E6E6E6", // 34 Light Gray  
                "#B3B3FF", // 35 Periwinkle  
                "#FFB3B3", // 36 Light Pink  
                "#B3FFE6", // 37 Mint  
                "#FFB3FF", // 38 Light Purple  
                "#B3B3FF", // 39 Periwinkle  
                "#FFCC99", // 40 Light Orange  
                "#99FFCC", // 41 Light Teal  
                "#CC99FF", // 42 Light Violet  
                "#FFFF99", // 43 Light Yellow  
                "#99FFFF", // 44 Light Cyan  
                "#FF99CC", // 45 Light Magenta  
                "#CCFF99", // 46 Light Lime  
                "#99CCFF", // 47 Light Sky Blue  
                "#FF6666", // 48 Light Coral  
                "#66FF99", // 49 Light Mint  
                "#6699FF", // 50 Light Azure  
                "#FF99FF", // 51 Light Fuchsia  
                "#99FF66", // 52 Light Chartreuse  
                "#66FFFF", // 53 Light Cyan  
                "#FFCC66", // 54 Light Gold  
                "#CCCCCC", // 55 Dark Gray  
                "#FF6600", // 56 Deep Orange  
                // Add more mappings as needed up to 56  
            };

            if (indexedColor >= 0 && indexedColor < indexedColors.Length)
            {
                return indexedColors[indexedColor];
            }

            return "None"; // Fallback for undefined indices  
        }

        /// <summary>  
        /// Converts a cell reference (e.g., "A1") to a one-based column index.  
        /// </summary>  
        private static int GetColumnIndexFromCellReference(string cellReference)
        {
            // Extract the column letters from the cell reference  
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", "");
            int columnIndex = 0;
            int factor = 1;

            for (int i = columnReference.Length - 1; i >= 0; i--)
            {
                columnIndex += (columnReference[i] - 'A' + 1) * factor;
                factor *= 26;
            }

            return columnIndex;
        }

        /// <summary>  
        /// Retrieves the cell value, handling shared strings, inline strings, and boolean types.  
        /// </summary>  
        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null || cell.CellValue == null)
                return string.Empty;

            string value = cell.CellValue.InnerText;

            if (cell.DataType != null)
            {
                // Fully qualify the CellValues enum to avoid namespace conflicts  
                CellValues dataType = cell.DataType.Value;

                if (dataType == CellValues.SharedString)
                {
                    var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                    if (sharedStringTable != null && int.TryParse(value, out int sharedStringIndex))
                    {
                        var sharedStringItem = sharedStringTable.ElementAtOrDefault(sharedStringIndex);
                        return sharedStringItem?.InnerText ?? string.Empty;
                    }
                }
                else if (dataType == CellValues.Boolean)
                {
                    return value == "0" ? "FALSE" : "TRUE";
                }
                else if (dataType == CellValues.InlineString)
                {
                    return cell.InlineString?.Text?.Text ?? string.Empty;
                }
                // Add additional else-if blocks for other CellValues types as needed  
            }

            // If the cell doesn't have a DataType or it's of a type we're not handling, return the raw value  
            return value;
        }

        /// <summary>  
        /// Creates a map of all cells in merged ranges to their merged value.  
        /// </summary>  
        private static Dictionary<string, string> CreateMergedCellMap(WorksheetPart worksheetPart, WorkbookPart workbookPart, IEnumerable<MergeCell> mergeCells)
        {
            var mergedCellMap = new Dictionary<string, string>();

            foreach (var mergeCell in mergeCells)
            {
                string mergeRange = mergeCell.Reference;
                string[] parts = mergeRange.Split(':'); // Example: "A1:B2"  

                if (parts.Length != 2)
                    continue; // Invalid merge range  

                string startCellRef = parts[0];
                string endCellRef = parts[1];

                // Get the value of the start cell  
                Cell startCell = GetCell(worksheetPart, startCellRef);
                if (startCell == null)
                    continue; // Start cell doesn't exist  

                string mergedValue = GetCellValue(startCell, workbookPart);

                // Get all cell references within the merge range  
                var cellRefs = GetCellsInRange(startCellRef, endCellRef);

                foreach (var cellRef in cellRefs)
                {
                    // Assign the merged value to each cell in the range  
                    mergedCellMap[cellRef] = mergedValue;
                }
            }

            return mergedCellMap;
        }

        /// <summary>  
        /// Retrieves a cell from the worksheet based on its reference.  
        /// </summary>  
        private static Cell GetCell(WorksheetPart worksheetPart, string cellReference)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Extract the row number  
            uint rowNumber = GetRowIndex(cellReference);
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowNumber);
            if (row == null)
                return null;

            return row.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference.Value, cellReference, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>  
        /// Generates all cell references within a specified range.  
        /// </summary>  
        private static List<string> GetCellsInRange(string start, string end)
        {
            var cells = new List<string>();

            var startCol = GetColumnName(start);
            var startRow = GetRowIndex(start);
            var endCol = GetColumnName(end);
            var endRow = GetRowIndex(end);

            int startColIndex = GetColumnIndexFromColumnName(startCol);
            int endColIndex = GetColumnIndexFromColumnName(endCol);

            for (int row = (int)startRow; row <= (int)endRow; row++)
            {
                for (int col = startColIndex; col <= endColIndex; col++)
                {
                    string columnName = GetColumnName(col);
                    string cellRef = columnName + row;
                    cells.Add(cellRef);
                }
            }

            return cells;
        }

        /// <summary>  
        /// Extracts the column name from a cell reference (e.g., "B" from "B2").  
        /// </summary>  
        private static string GetColumnName(string cellReference)
        {
            return Regex.Replace(cellReference.ToUpper(), @"[\d]", "");
        }

        /// <summary>  
        /// Extracts the row index from a cell reference (e.g., "2" from "B2").  
        /// </summary>  
        private static uint GetRowIndex(string cellReference)
        {
            string rowPart = Regex.Replace(cellReference.ToUpper(), @"[A-Z]", "");
            return uint.TryParse(rowPart, out uint row) ? row : 0;
        }

        /// <summary>  
        /// Converts a column name (e.g., "B") to a one-based index (e.g., 2).  
        /// </summary>  
        private static int GetColumnIndexFromColumnName(string columnName)
        {
            int columnIndex = 0;
            int factor = 1;

            for (int i = columnName.Length - 1; i >= 0; i--)
            {
                columnIndex += (columnName[i] - 'A' + 1) * factor;
                factor *= 26;
            }

            return columnIndex;
        }

        /// <summary>  
        /// Converts a one-based column index (e.g., 2) to a column name (e.g., "B").  
        /// </summary>  
        private static string GetColumnName(int columnIndex)
        {
            string columnName = string.Empty;
            while (columnIndex > 0)
            {
                int modulo = (columnIndex - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                columnIndex = (columnIndex - modulo) / 26;
            }
            return columnName;
        }

        /// <summary>  
        /// Represents a table extracted from the worksheet.  
        /// </summary>  
        public class Table
        {
            [Required]
            public Guid GUID { get; set; }

            public List<SheetRow> Rows { get; set; }

            public int SheetIndex { get; set; }
        }
    }
}