using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Drawing;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OutSystems.ExternalLibraries.SDK;

namespace OutSystems.ExternalLib.Excel;

public class ExcelLibrary : IExcelLibrary
{
    private int columnNumber(String columnName)
    {
        char[] chars = columnName.ToUpper().ToCharArray();

        return (int)(Math.Pow(26, chars.Count() - 1)) * 
            (System.Convert.ToInt32(chars[0]) - 64) + 
            ((chars.Count() > 2) ? columnNumber(columnName.Substring(1, columnName.Length - 1)) : 
            ((chars.Count() == 2) ? (System.Convert.ToInt32(chars[chars.Count() - 1]) - 64) : 0));
    }

    private string[] SplitRegex(string inputStr) {
        string pattern = @"^([a-zA-Z]+)([0-9]+)$";
        Regex rgx = new Regex(pattern);
        //string input = "A4";
        string[] result = rgx.Split(inputStr).Where(s => s != String.Empty).ToArray<string>();
        return result;       
    }

    private bool isValidHexColor(string hexColor) {
        var r = new Regex("^#[abcdefABCDEF0-9]{6}$");
        return r.IsMatch(hexColor);
    }

    public ExcelPackage Excel_Open(byte[] excelBinary)
    {
        Stream stream = new MemoryStream(excelBinary);
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        ExcelTextFormat format = new ExcelTextFormat();
        format.Encoding = new System.Text.UTF8Encoding();

        ExcelPackage? package = null;
        try
        {
            package = new ExcelPackage(stream);
        }
        catch(Exception exception)
        {
            // Added check Valid Excel File in 30 January 2024
            throw new System.Exception("Excel File Invalid Format!", exception);
        }

        return package;
    }

    private ExcelWorksheet Worksheet_Select(ExcelPackage package, string? sheetName)
    {
        if (sheetName == null || sheetName == "") return package.Workbook.Worksheets[0];

        ExcelWorksheet? worksheet = package.Workbook.Worksheets[sheetName] ?? throw new System.NullReferenceException("Can't find Sheet Name: " + sheetName + "!");
        return worksheet;
    }

    private byte[] AddSheets(ExcelPackage package, Worksheet[] worksheets) {
        ExcelWorksheet excelWorksheet;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (package)
        {
            if(worksheets.Length > 0) {
                foreach(Worksheet worksheet in worksheets) {
                    if(worksheet.Name != null) {
                        string sheetName = worksheet.Name;
                        if(sheetName == "") sheetName = Next_SheetName(package);
                        excelWorksheet = package.Workbook.Worksheets.Add(sheetName);
                        // var r = new Regex("^#[abcdefABCDEF0-9]{6}$");
                        if(worksheet.ColorHex != "" || worksheet.ColorHex != null) {
                            if(isValidHexColor(worksheet.ColorHex!)) {
                                excelWorksheet.TabColor = ColorTranslator.FromHtml(worksheet.ColorHex!);
                            }
                        }
                    }
                }
            } else {
                string sheetName = "";
                sheetName = Next_SheetName(package);
                package.Workbook.Worksheets.Add(sheetName);
            }
            return package.GetAsByteArray();
        }
    }

    private string Next_SheetName(ExcelPackage package)
    {
        string sheetName = "";
        int c = package.Workbook.Worksheets.Count + 1;
        sheetName = "Sheet " + c;
        return sheetName;
    }

    private ExcelRange? Cell_Selections(ExcelPackage package, int StartRowNumber, int StartColNumber, int EndRowNumber, int EndColNumber, string CellName, string? sheetName = null) {
        // Console.WriteLine("Start Row: " + StartRowNumber);
        // Console.WriteLine("Start Col: " + StartColNumber);
        // Console.WriteLine("End Row: " + EndRowNumber);
        // Console.WriteLine("End Col: " + EndColNumber);
        // Console.WriteLine("Cell Name: " + CellName);

        if(StartRowNumber > 0 && StartColNumber > 0 && EndRowNumber <= 0 && EndColNumber <= 0) {
            // Console.WriteLine("Cell");
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
            return worksheet.Cells[StartRowNumber, StartColNumber];
        }

        if(StartRowNumber > 0 && EndRowNumber > 0 && StartRowNumber <= EndRowNumber && StartColNumber <= EndColNumber) {
            // Console.WriteLine("Range");
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
            ExcelRange excelRange = worksheet.Cells[StartRowNumber, StartColNumber, EndRowNumber, EndColNumber];
            return excelRange;
        }

        if(CellName != null && CellName != "") {
            Console.WriteLine("Name: " + CellName);
            var r = new Regex("^([a-zA-Z]{1,7})([0-9]+)$");
            if(r.IsMatch(CellName)) {
                Console.WriteLine("Cell Name Match with Regex Single Cell");
                ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
                return worksheet.Cells[CellName];
            } else {
                var r2 = new Regex("^([a-zA-Z]{1,7})([0-9]{1,7})[:]([a-zA-Z]{1,7})([0-9]{1,7})$");
                if(r2.IsMatch(CellName)) {
                    Console.WriteLine("Cell Name Match with Regex Multi Cell");
                    ExcelWorksheet worksheet2 = Worksheet_Select(package, sheetName);
                    return worksheet2.Cells[CellName];
                } else {
                    Console.WriteLine("Cell Name NOT Match with Regex");

                    var r3 = new Regex("^([a-zA-Z]{1,7})[:]([a-zA-Z]{1,7})$");
                    if(r3.IsMatch(CellName)) {
                        Console.WriteLine("Cell Name Match with Regex Multi Cell - 3");
                        ExcelWorksheet worksheet3 = Worksheet_Select(package, sheetName);
                        ExcelRange excelRange = worksheet3.Cells[CellName];

                        Console.WriteLine("Sheet Name: " + excelRange.Worksheet.Name);
                        Console.WriteLine("Start Row: " + excelRange.Start.Row);
                        Console.WriteLine("Start Col: " + excelRange.Start.Column);
                        Console.WriteLine("End Row: " + excelRange.End.Row);
                        Console.WriteLine("End Col: " + excelRange.End.Column);

                        return excelRange;
                    } else {
                        Console.WriteLine("Cell Name Match with Regex Multi Cell - 4");
                        var r4 = new Regex("^([a-zA-Z]{1,7})$");
                        if(r4.IsMatch(CellName)) {
                            Console.WriteLine("Single Column");
                            CellName = CellName + ":" + CellName;
                            ExcelWorksheet worksheet3 = Worksheet_Select(package, sheetName);
                            ExcelRange excelRange = worksheet3.Cells[CellName];

                            Console.WriteLine("Sheet Name: " + excelRange.Worksheet.Name);
                            Console.WriteLine("Start Row: " + excelRange.Start.Row);
                            Console.WriteLine("Start Col: " + excelRange.Start.Column);
                            Console.WriteLine("End Row: " + excelRange.End.Row);
                            Console.WriteLine("End Col: " + excelRange.End.Column);

                            return excelRange;
                        } else {
                            var excelNamedRange = package.Workbook.Names[CellName];
                            ExcelWorksheet worksheet = Worksheet_Select(package, excelNamedRange.Worksheet.Name);
                            ExcelRange excelRange = worksheet.Cells[excelNamedRange.Start.Row, excelNamedRange.Start.Column, excelNamedRange.End.Row, excelNamedRange.End.Column];

                            Console.WriteLine("Sheet Name: " + excelNamedRange.Worksheet.Name);
                            Console.WriteLine("Start Row: " + excelNamedRange.Start.Row);
                            Console.WriteLine("Start Col: " + excelNamedRange.Start.Column);
                            Console.WriteLine("End Row: " + excelNamedRange.End.Row);
                            Console.WriteLine("End Col: " + excelNamedRange.End.Column);
                            return excelRange;
                        }
                    }
                }                
            }
        }

        if(StartRowNumber == 0 || StartColNumber == 0) {
            throw new System.NullReferenceException("Start Row or Cell, must be greater than 0");
        }

        throw new System.NullReferenceException("Cell Selections Failed!");
    }

    private ExcelWorksheet Border_Format(ExcelRange cellRange, BorderStyleFormat borderStyleFormat) {


        ExcelBorderStyle excelBorderStyle;
        System.Drawing.Color colorVar = ColorTranslator.FromHtml("#000000");

        if(borderStyleFormat.BorderColorHex != "") {
            if(isValidHexColor(borderStyleFormat.BorderColorHex)) {
                colorVar = ColorTranslator.FromHtml(borderStyleFormat.BorderColorHex);
            }
        }

        if (Enum.TryParse<ExcelBorderStyle>(borderStyleFormat.BorderStyle, out excelBorderStyle))
        {
            if(borderStyleFormat.IsRound) {
                cellRange.Style.Border.BorderAround(excelBorderStyle, colorVar);
            } 

            if(borderStyleFormat.IsTop) {
                cellRange.Style.Border.Top.Style = excelBorderStyle;
                cellRange.Style.Border.Top.Color.SetColor(colorVar);
            }

            if(borderStyleFormat.IsBottom) {
                cellRange.Style.Border.Bottom.Style = excelBorderStyle;
                cellRange.Style.Border.Bottom.Color.SetColor(colorVar);
            }

            if(borderStyleFormat.IsLeft) {
                cellRange.Style.Border.Left.Style = excelBorderStyle;
                cellRange.Style.Border.Left.Color.SetColor(colorVar);
            }

            if(borderStyleFormat.IsRight) {
                cellRange.Style.Border.Right.Style = excelBorderStyle;
                cellRange.Style.Border.Right.Color.SetColor(colorVar);
            }
        }

        return cellRange.Worksheet;
    }


    private ExcelWorksheet Cell_Format(ExcelRange cellRange, CellFormat cellFormat)
    {
        cellRange.Style.Numberformat.Format = cellFormat.CellTypeFormat;

        if (cellFormat.BackgroundColorHex != "")
        {
            cellRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cellRange.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(cellFormat.BackgroundColorHex));
        }

        if (cellFormat.FontStyleFormat.FontColorHex != "")
        {
            cellRange.Style.Font.Color.SetColor(ColorTranslator.FromHtml(cellFormat.FontStyleFormat.FontColorHex));
        }

        if (cellFormat.FontStyleFormat.FontSize > 0)
        {
            cellRange.Style.Font.Size = cellFormat.FontStyleFormat.FontSize;
        }

        cellRange.Style.Font.Bold = cellFormat.FontStyleFormat.IsBold;
        cellRange.Style.Font.Italic = cellFormat.FontStyleFormat.IsItalic;
        cellRange.Style.Font.UnderLine = cellFormat.FontStyleFormat.IsUnderline;

        cellRange.Style.ShrinkToFit = cellFormat.FontStyleFormat.IsShrinkToFit;
        cellRange.Style.WrapText = cellFormat.FontStyleFormat.IsWrapText;

        cellRange.Style.QuotePrefix = cellFormat.FontStyleFormat.IsQuotePrefix;

        cellRange.Style.Locked = cellFormat.IsLocked;
        cellRange.Style.Hidden = cellFormat.IsHidden;

        if(cellFormat.IsAutoFitColumn) {
            
            cellRange.Worksheet.Columns[cellRange.Start.Column].AutoFit();
            //cellRange.AutoFitColumns();
        }

        if(cellFormat.FontStyleFormat.HorizontalAlignment.ToLower() == "center") cellRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        if(cellFormat.FontStyleFormat.HorizontalAlignment.ToLower() == "left") cellRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        if(cellFormat.FontStyleFormat.HorizontalAlignment.ToLower() == "right") cellRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

        if(cellFormat.FontStyleFormat.VerticalAlignment.ToLower() == "top") cellRange.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
        if(cellFormat.FontStyleFormat.VerticalAlignment.ToLower() == "center") cellRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        if(cellFormat.FontStyleFormat.VerticalAlignment.ToLower() == "bottom") cellRange.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

        return cellRange.Worksheet;
    }

    private string GenerateGUID()
    {
        Guid g = Guid.NewGuid();
        return g.ToString();
    }


// ============================================================
// Public Method Implementation Interface - WORKBOOK
// ============================================================

    public byte[] Workbook_Create(Worksheet[] worksheets)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        return AddSheets(new ExcelPackage(), worksheets);
    }

    public Worksheet[] Workbook_GetWorksheet(byte[] excelBinary)
    {
        using (var package = Excel_Open(excelBinary))
        {
            int countWS = package.Workbook.Worksheets.Count;
            Worksheet[] worksheets = new Worksheet[countWS];
            for(int i=0;i<countWS;i++) {
                worksheets[i].Index = i;
                worksheets[i].Name = package.Workbook.Worksheets[i].Name;
                worksheets[i].ColorHex = ColorTranslator.ToHtml(package.Workbook.Worksheets[i].TabColor);
            }
            return worksheets;
        }
    }

    public byte[] Workbook_SetProperties(byte[] excelBinary, WorkbookProperties workbookProperties)
    {
        using (var package = Excel_Open(excelBinary))
        {
            package.Workbook.Properties.Title = workbookProperties.Title ?? "";        
            package.Workbook.Properties.Author = workbookProperties.Author ?? "";        
            package.Workbook.Properties.Comments = workbookProperties.Comments ?? "";        
            package.Workbook.Properties.Company = workbookProperties.Company ?? "";
            package.Workbook.Properties.Subject = workbookProperties.Subject ?? "";
            package.Workbook.Properties.Manager = workbookProperties.Manager ?? "";
            package.Workbook.Properties.Category = workbookProperties.Category ?? "";
            package.Workbook.Properties.Keywords = workbookProperties.Keywords ?? "";
            
            if(workbookProperties.KeyValues == null) {
                return package.GetAsByteArray();
            }

            foreach(KeyValue keyValue in workbookProperties.KeyValues) {
                package.Workbook.Properties.SetCustomPropertyValue(keyValue.Key, keyValue.Value);
            }

            return package.GetAsByteArray();
        }    
    }

// ============================================================
// Public Method Implementation Interface - WORKSHEET
// ============================================================

    public byte[] Worksheet_Add(byte[] excelBinary, string? sheetName = null)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        if(sheetName == null) {
            return AddSheets(Excel_Open(excelBinary), new Worksheet[0]);
        } else {
            Worksheet[] worksheets = new Worksheet[1];
            worksheets[0].Name = sheetName;
            return AddSheets(Excel_Open(excelBinary), worksheets);
        }
    }

    public byte[] Worksheet_AddList(byte[] excelBinary, Worksheet[] worksheets)
    {
        if (worksheets.Length == 0) return excelBinary;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        return AddSheets(Excel_Open(excelBinary), worksheets);
    }

    public byte[] Worksheet_AutofitColumns(byte[] excelBinary, string? sheetName = null)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
            worksheet.Cells.AutoFitColumns();
            return package.GetAsByteArray();
        }
    }

    public byte[] Worksheet_Calculate(byte[] excelBinary, string? sheetName = null)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
            worksheet.Calculate();
            return package.GetAsByteArray();
        }
    }   

    public byte[] Worksheet_Protect(byte[] excelBinary, string password, bool? isAllowAutoFilter = false, bool? isAllowDeleteColumns = false, bool? isAllowDeleteRows = false, 
        bool? isAllowEditObject = false, bool? isAllowFormatCells = false, bool? isAllowFormatColumns = false, bool? isAllowFormatRows = false, bool? isAllowInsertColumns = false, 
        bool? isAllowInsertHyperlinks = false, bool? isAllowInsertRows = false, bool? isAllowPivotTables = false, bool? isAllowSelectLockedCells = false, bool? isAllowSelectUnLockedCells = false, 
        bool? isAllowSort = false, bool? isProtected = false, string? sheetName = null)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);

            worksheet.Protection.AllowAutoFilter = isAllowAutoFilter.GetValueOrDefault();
            worksheet.Protection.AllowDeleteColumns = isAllowDeleteColumns.GetValueOrDefault();
            worksheet.Protection.AllowDeleteRows = isAllowDeleteRows.GetValueOrDefault();

            worksheet.Protection.AllowEditObject = isAllowEditObject.GetValueOrDefault();

            worksheet.Protection.AllowFormatCells = isAllowFormatCells.GetValueOrDefault();
            worksheet.Protection.AllowFormatColumns = isAllowFormatColumns.GetValueOrDefault();
            worksheet.Protection.AllowFormatRows = isAllowFormatRows.GetValueOrDefault();

            worksheet.Protection.AllowInsertColumns = isAllowInsertColumns.GetValueOrDefault();
            worksheet.Protection.AllowInsertHyperlinks = isAllowInsertHyperlinks.GetValueOrDefault();
            worksheet.Protection.AllowInsertRows = isAllowInsertRows.GetValueOrDefault();

            worksheet.Protection.AllowPivotTables = isAllowPivotTables.GetValueOrDefault();
            worksheet.Protection.AllowSelectLockedCells = isAllowSelectLockedCells.GetValueOrDefault();
            worksheet.Protection.AllowSelectUnlockedCells = isAllowSelectUnLockedCells.GetValueOrDefault();
            worksheet.Protection.AllowSort = isAllowSort.GetValueOrDefault();

            
            worksheet.Protection.IsProtected = isProtected.GetValueOrDefault();

            worksheet.Protection.SetPassword(password);

            return package.GetAsByteArray();
        }        
    }

    public byte[] Worksheet_AddAutoFilter(byte[] excelBinary, int startCellRow = 0, int startCellColumn = 0, int endCellRow = 0, int endCellColumn = 0, string? cellName = null, string? sheetName = null)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelRange? excelRanges;

            try {
                excelRanges = Cell_Selections(package, startCellRow, startCellColumn, endCellRow, endCellColumn, cellName ?? "", sheetName);
            } catch {
                return package.GetAsByteArray();
            }

            if(excelRanges == null) {
                return package.GetAsByteArray();
            }

            excelRanges.AutoFilter = true;
            return package.GetAsByteArray();
        }
    }

    public byte[] Worksheet_Delete(byte[] excelBinary, int sheetIndex = 0, string? sheetName = null)
    {
        using (var package = Excel_Open(excelBinary))
        {
            if(sheetName == null) {
                package.Workbook.Worksheets.Delete(sheetIndex);
            } else {
                ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
                package.Workbook.Worksheets.Delete(worksheet.Index);
            }
            return package.GetAsByteArray();        
        }
    }

    public byte[] Worksheet_Rename(byte[] excelBinary, string newSheetName, int sheetIndex = 0, string? sheetName = null)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet;
            if(sheetName == null) {
                package.Workbook.Worksheets[sheetIndex].Name = newSheetName;
            } else {
                worksheet = Worksheet_Select(package, sheetName);
                worksheet.Name = newSheetName;
            }

            return package.GetAsByteArray();
        }   
    }

    public byte[] Worksheet_Hide_Show(byte[] excelBinary, int sheetIndex = 0, string? sheetName = null, bool isShow = false)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet;
            if(sheetName == null) {
                worksheet = package.Workbook.Worksheets[sheetIndex];
            } else {
                worksheet = Worksheet_Select(package, sheetName);
            }
            if(isShow) {
                worksheet.Hidden = OfficeOpenXml.eWorkSheetHidden.Visible;
            } else {
                worksheet.Hidden = OfficeOpenXml.eWorkSheetHidden.Hidden;
            }
            return package.GetAsByteArray();
        }        
    }

// ============================================================
// Public Method Implementation Interface - CELL
// ============================================================

    public string Cell_Read(byte[] excelBinary, int cellRow = 0, int cellColumn = 0, string? cellName = null, string? sheetName = null)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelRange? excelRanges;
            try {
                excelRanges = Cell_Selections(package, cellRow, cellColumn, 0, 0, cellName ?? "", sheetName);
            } catch {
                return "";
            }
            
            if(excelRanges == null) {
                return "";
            }

            string? cellValue = Convert.ToString(excelRanges.Value);
            if (cellValue == null) return "";
            
            return cellValue;
        }
    }

    public byte[] Cell_Merge(byte[] excelBinary, CellMerge[] cellMerges)
    {
        using (var package = Excel_Open(excelBinary))
        {
            foreach(CellMerge cellMerge in cellMerges) {
                ExcelRange? excelRanges;
                try {
                    excelRanges = Cell_Selections(package, cellMerge.CellRange.StartCellRow, cellMerge.CellRange.StartCellColumn, cellMerge.CellRange.EndCellRow, cellMerge.CellRange.EndCellColumn, cellMerge.CellName ?? "", cellMerge.SheetName);

                } catch {
                    return package.GetAsByteArray();
                }

                if(excelRanges == null) {
                    return package.GetAsByteArray();
                }

                excelRanges.Merge = true;
            }

            return package.GetAsByteArray();
        }
    }

    public byte[] Cell_UnMerge(byte[] excelBinary, CellMerge[] cellUnMerges)
    {
        using (var package = Excel_Open(excelBinary))
        {
            foreach(CellMerge cellUnMerge in cellUnMerges) {
                ExcelRange? excelRanges;
                try {
                    excelRanges = Cell_Selections(package, cellUnMerge.CellRange.StartCellRow, cellUnMerge.CellRange.StartCellColumn, cellUnMerge.CellRange.EndCellRow, cellUnMerge.CellRange.EndCellColumn, cellUnMerge.CellName ?? "", cellUnMerge.SheetName);

                } catch {
                    return package.GetAsByteArray();
                }

                if(excelRanges == null) {
                    return package.GetAsByteArray();
                }

                excelRanges.Merge = false;
            }

            return package.GetAsByteArray();
        }
    }

    public byte[] Cell_Write(byte[] excelBinary, CellWrite[] cellWrites, CellCopy? cellCopy = null)
    {
        using (var package = Excel_Open(excelBinary))
        {
            foreach(CellWrite cellWrite in cellWrites) {
                ExcelWorksheet worksheet = Worksheet_Select(package, cellWrite.SheetName);

                ExcelRange? excelRanges;
                try {
                    excelRanges = Cell_Selections(package, cellWrite.Cell.CellRow, cellWrite.Cell.CellColumn, 0, 0, cellWrite.CellName ?? "", cellWrite.SheetName);

                } catch {
                    return package.GetAsByteArray();
                } 
                
                if(excelRanges == null) {
                    return package.GetAsByteArray();
                }

                if (cellWrite.CellFormat.CellType.ToLower() == "number")
                {
                    excelRanges.Value = Convert.ToInt64(cellWrite.CellValue);
                }
                else if (cellWrite.CellFormat.CellType.ToLower() == "decimal")
                {
                    excelRanges.Value = Convert.ToDecimal(cellWrite.CellValue);
                }
                else if (cellWrite.CellFormat.CellType.ToLower() == "datetime" || cellWrite.CellFormat.CellType.ToLower() == "date")
                {
                    excelRanges.Value = Convert.ChangeType(cellWrite.CellValue, typeof(DateTime));
                }
                else if (cellWrite.CellFormat.CellType.ToLower() == "bool")
                {
                }
                else if (cellWrite.CellFormat.CellType.ToLower() == "formula")
                {
                    excelRanges.Formula = cellWrite.CellValue;
                }
                else
                {
                    excelRanges.Value = cellWrite.CellValue;
                }

                worksheet = Cell_Format(excelRanges, cellWrite.CellFormat);

            }

            // Console.WriteLine("Row: " + cellCopy.GetValueOrDefault().SourceCell.CellRow);
            // Console.WriteLine("Col: " + cellCopy.GetValueOrDefault().SourceCell.CellColumn);
            // Console.WriteLine("Source Name: " + cellCopy.GetValueOrDefault().SourceCellName);
            // Console.WriteLine("StartRow: " + cellCopy.GetValueOrDefault().DestinationCell.StartCellRow);
            // Console.WriteLine("StartCol: " + cellCopy.GetValueOrDefault().DestinationCell.StartCellColumn);
            // Console.WriteLine("EndRow: " + cellCopy.GetValueOrDefault().DestinationCell.EndCellRow);
            // Console.WriteLine("EndCol: " + cellCopy.GetValueOrDefault().DestinationCell.EndCellColumn);
            // Console.WriteLine("Dest Name: " + cellCopy.GetValueOrDefault().DestinationCellName);
            // Console.WriteLine("IsEmpty: " + cellCopy.GetValueOrDefault().IsEmpty());

            if(!cellCopy.GetValueOrDefault().IsEmpty()) {
                return Cell_Copy(package.GetAsByteArray(), cellCopy.GetValueOrDefault());
            }

            return package.GetAsByteArray();
        }        
    }

    public byte[] Cell_Copy(byte[] excelBinary, CellCopy cellCopy)
    {
        using (var package = Excel_Open(excelBinary))
        {

            ExcelWorksheet worksheet = Worksheet_Select(package, cellCopy.SheetName);
            ExcelRange? sourceExcelRange = Cell_Selections(package, cellCopy.SourceCell.CellRow, cellCopy.SourceCell.CellColumn, 0,0, cellCopy.SourceCellName, cellCopy.SheetName);

            if(sourceExcelRange == null) {
                return package.GetAsByteArray();
            }

            int sourceRow = sourceExcelRange.Start.Row;
            int sourceCol = sourceExcelRange.Start.Column;

            ExcelRange? destExcelRange = Cell_Selections(package, cellCopy.DestinationCell.StartCellRow, cellCopy.DestinationCell.StartCellColumn, cellCopy.DestinationCell.EndCellRow, cellCopy.DestinationCell.EndCellColumn, cellCopy.DestinationCellName, cellCopy.SheetName);

            if(destExcelRange == null) {
                return package.GetAsByteArray();
            }

            int destStartRow = destExcelRange.Start.Row;
            int destStartCol = destExcelRange.Start.Column;
            int destEndRow = destExcelRange.End.Row;
            int destEndCol = destExcelRange.End.Column;

            int countRow = destStartRow;
            int countCol = destStartCol;

            while (countRow <= destEndRow) {
                while (countCol <= destEndCol) {
                    sourceExcelRange.Copy(worksheet.Cells[countRow, countCol]);
                    countCol++;
                }
                countCol = destStartCol;
                countRow++;
            }

            return package.GetAsByteArray();        
        }
    }

    public CellFindResult[] Cell_FindByValue(byte[] excelBinary, string cellValue, bool isContain = false, string? cellRange = null, string? sheetName = null)
    {
        using (var package = Excel_Open(excelBinary))
        {
            List<CellFindResult> cellsList = new List<CellFindResult>();
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);

            if(cellRange == null || cellRange == "") {
                int rowStart = worksheet.Dimension.Start.Row; 
                int rowEnd = worksheet.Dimension.End.Row;
                cellRange = rowStart.ToString() + ":" + rowEnd.ToString();
            }

            IEnumerable<ExcelRangeBase> searchCell;

            if(isContain == true) {
                searchCell = from cell in worksheet.Cells[cellRange]
                    where cell.Value != null && cell.Value?.ToString()!.Contains(cellValue) == true
                    select cell;
            } else {
                searchCell = from cell in worksheet.Cells[cellRange]
                    where cell.Value != null && cell.Value?.ToString() == cellValue
                    select cell;
            }


            foreach (ExcelRangeBase cellResult in searchCell) {
                Cell cellItem = new Cell();
                cellItem.CellRow = cellResult.Start.Row;
                cellItem.CellColumn = cellResult.Start.Column;

                CellFindResult cellFindResult = new CellFindResult();
                cellFindResult.Cell = cellItem;
                cellFindResult.CellName = cellResult.Start.Address;
                cellFindResult.CellValue = cellResult.Value.ToString()!;

                cellsList.Add(cellFindResult);
            }

            return cellsList.ToArray<CellFindResult>();
        }     
    }

    public byte[] Cell_Write_RichText(byte[] excelBinary, CellWriteRichText[] cellWriteRichTexts) {
        using (var package = Excel_Open(excelBinary))
        {
            foreach(CellWriteRichText cellWriteRichText in cellWriteRichTexts) {
                ExcelWorksheet worksheet = Worksheet_Select(package, cellWriteRichText.SheetName);
                ExcelRange? excelRanges;

                try {
                    excelRanges = Cell_Selections(package, cellWriteRichText.Cell.CellRow, cellWriteRichText.Cell.CellColumn, 0, 0, cellWriteRichText.CellName ?? "", cellWriteRichText.SheetName);
                } catch {
                    return package.GetAsByteArray();
                }
                
                if(excelRanges == null) {
                    return package.GetAsByteArray();
                }

                excelRanges.IsRichText = true;
                // Console.WriteLine("Range: " + excelRanges.Address);

                foreach(RichTextFormatText richTextFormatText in cellWriteRichText.RichTextFormatTexts) {
                    var richFormatCell = excelRanges.RichText.Add(richTextFormatText.CellValue);  
                    if(richTextFormatText.FontName != "") richFormatCell.FontName = richTextFormatText.FontName; else richFormatCell.FontName = "Calibri";
                    if(richTextFormatText.FontSize > 0) richFormatCell.Size = richTextFormatText.FontSize; else richFormatCell.Size = 11;
                    if(richTextFormatText.FontColorHex != "") richFormatCell.Color = ColorTranslator.FromHtml(richTextFormatText.FontColorHex); else richFormatCell.Color = System.Drawing.Color.Black;
                    richFormatCell.Bold = richTextFormatText.IsBold;
                    richFormatCell.Italic = richTextFormatText.IsItalic;
                    richFormatCell.UnderLine = richTextFormatText.IsUnderline;
                    richFormatCell.Strike = richTextFormatText.IsStrikeOut;
                }

                if(cellWriteRichText.IsAutoFitColumn) {
                    excelRanges.Worksheet.Columns[excelRanges.Start.Column].AutoFit();
                }
                
                if(cellWriteRichText.HorizontalAlignment.ToLower() == "center") excelRanges.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                if(cellWriteRichText.HorizontalAlignment.ToLower() == "left") excelRanges.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                if(cellWriteRichText.HorizontalAlignment.ToLower() == "right") excelRanges.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                if(cellWriteRichText.VerticalAlignment.ToLower() == "top") excelRanges.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                if(cellWriteRichText.VerticalAlignment.ToLower() == "center") excelRanges.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                if(cellWriteRichText.VerticalAlignment.ToLower() == "bottom") excelRanges.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

                // Console.WriteLine("R: " + excelRanges.Value);
            }
            return package.GetAsByteArray();
        }
    }

// ============================================================
// Public Method Implementation Interface - COLUMN
// ============================================================

    public byte[] Column_Delete(byte[] excelBinary, int colIndex, string? sheetName = null)
    {
        if(colIndex < 1) return excelBinary;        
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
            worksheet.DeleteColumn(colIndex);
            return package.GetAsByteArray();
        }
    }

    public byte[] Column_Hide_Show(byte[] excelBinary, int colIndex, bool isShow = false, string? sheetName = null)
    {
        if(colIndex < 1) return excelBinary;        
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
            worksheet.Columns[colIndex].Hidden = isShow;
            return package.GetAsByteArray();
        }
    }

    public byte[] Column_Insert(byte[] excelBinary, int colIndex, int colNewAdd = 1, int colWidth = 64, bool isCopyFormatFromSource = false, string? sheetName = null)
    {
        if(colIndex < 1) return excelBinary;        
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);

            if(isCopyFormatFromSource) {
                worksheet.InsertColumn(colIndex, colNewAdd, colIndex);
            } else {
                worksheet.InsertColumn(colIndex, colNewAdd);
            }            
            
            int endCol = colIndex + colNewAdd;
            if(colWidth != 64) {
                for(int cIndex = colIndex;cIndex <= endCol; cIndex++) {
                    worksheet.Columns[cIndex].Width = colWidth;
                }
            }

            return package.GetAsByteArray();
        }
    }

    public byte[] Column_Width(byte[] excelBinary, int colIndex, int colWidth = 64, string? sheetName = null)
    {
        if(colIndex < 1) return excelBinary;        
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
            worksheet.Columns[colIndex].Width = colWidth;
            return package.GetAsByteArray();
        }
    }    

// ============================================================
// Public Method Implementation Interface - ROW
// ============================================================

    public byte[] Row_Delete(byte[] excelBinary, int rowIndex, string? sheetName = null)
    {
        if(rowIndex < 1) return excelBinary;        
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
            worksheet.DeleteRow(rowIndex);
            return package.GetAsByteArray();
        }
    }

    public byte[] Row_Hide_Show(byte[] excelBinary, int rowIndex, bool isShow = false, string? sheetName = null)
    {
        if(rowIndex < 1) return excelBinary;        
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
            worksheet.Rows[rowIndex].Hidden = isShow;
            return package.GetAsByteArray();
        }
    }

    public byte[] Row_Insert(byte[] excelBinary, int rowIndex, int rowNewAdd = 1, int rowHeight = 20, bool isCopyFormatFromSource = false, string? sheetName = null)
    {
        if(rowIndex < 1) return excelBinary;        
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);

            if(isCopyFormatFromSource) {
                worksheet.InsertRow(rowIndex, rowNewAdd, rowIndex);
            } else {
                worksheet.InsertRow(rowIndex, rowNewAdd);
            }

            int endRow = rowIndex + rowNewAdd;
            if(rowHeight != 20) {
                for(int rIndex = rowIndex+1;rIndex <= endRow; rIndex++) {
                    worksheet.Rows[rowIndex].Height = rowHeight;
                }
            }
            return package.GetAsByteArray();
        }
    }

    public byte[] Row_Height(byte[] excelBinary, int rowIndex, int rowHeight = 20, string? sheetName = null)
    {
        if(rowIndex < 1) return excelBinary;
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);
            worksheet.Rows[rowIndex].Height = rowHeight;

            return package.GetAsByteArray();
        }

    }

// ============================================================
// Public Method Implementation Interface - Data Validations
// ============================================================

    private IExcelDataValidation Data_Validation_Config(ExcelWorksheet excelWorksheet, string Address, DataValidationConfig dataValidationConfig) {

        IExcelDataValidation validation = excelWorksheet.DataValidations[Address];

        validation.ShowErrorMessage = dataValidationConfig.IsShowErrorMessage;

        if (dataValidationConfig.ErrorStyle.ToLower() == "information")
        {
            validation.ErrorStyle = ExcelDataValidationWarningStyle.information;
        } else {
            if (dataValidationConfig.ErrorStyle.ToLower() == "stop")
            {
                validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            } else {
                validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
            }
        }

        validation.ErrorTitle = dataValidationConfig.ErrorTitle;
        validation.Error = dataValidationConfig.ErrorMessage;
        validation.ShowInputMessage = dataValidationConfig.IsShowInputMessage;
        validation.PromptTitle = dataValidationConfig.InputTitle;
        validation.Prompt = dataValidationConfig.InputMessage;

        return validation;
    }

    public byte[] Data_Validation_Integer(byte[] excelBinary, CellDataValidation cellDataValidation, DataValidation dataValidation) {
        using (var package = Excel_Open(excelBinary))
        {
            // Console.WriteLine(dataValidation.dataValidationConfig.ValidationOperator.ToLower());

            int ValA = 0;
            int ValB = 0;

            if(int.TryParse(dataValidation.FormulaValue1, out ValA)) {
                ValA = Convert.ToInt32(dataValidation.FormulaValue1);
            } else {
                throw new System.NullReferenceException("Formula 1 Value is not Integer!");
            }

            if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "between" || dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "not between") {
                if(int.TryParse(dataValidation.FormulaValue2, out ValB)) {
                    ValB = Convert.ToInt32(dataValidation.FormulaValue2);
                } else {
                    throw new System.NullReferenceException("Formula 2 Value is not Integer!");
                }

                // Console.WriteLine("Val A: " + ValA);
                // Console.WriteLine("Val B: " + ValB);

                if( ! (ValA >= 0 && ValB > 0 && ValB > ValA)) {
                    throw new System.NullReferenceException("Formula 1 and Formula 2 (Greater Than) Value required for between and not between Operator!");
                }

            }

            ExcelWorksheet worksheet = Worksheet_Select(package, cellDataValidation.SheetName);
            ExcelRange? targetExcelRange;
            try {
                targetExcelRange = Cell_Selections(package, cellDataValidation.CellRange.StartCellRow, cellDataValidation.CellRange.StartCellColumn, cellDataValidation.CellRange.EndCellRow, cellDataValidation.CellRange.EndCellColumn, cellDataValidation.CellName, cellDataValidation.SheetName);

            } catch {
                return package.GetAsByteArray();
            }

            if(targetExcelRange == null) {
                return package.GetAsByteArray();
            }

            worksheet.DataValidations.AddIntegerValidation(targetExcelRange.Address);
            ExcelDataValidationInt validation = (OfficeOpenXml.DataValidation.ExcelDataValidationInt) Data_Validation_Config(worksheet, targetExcelRange.Address, dataValidation.dataValidationConfig);

            // //Between (Default), Equal, GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, NotBetween, NotEqual

            if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "between" || dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "not between") { 
                validation.Operator = ExcelDataValidationOperator.between;
                validation.Formula.Value = ValA;
                validation.Formula2.Value = ValB;
            }  else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "equal") {
                    validation.Operator = ExcelDataValidationOperator.equal;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "greaterthan") {
                    validation.Operator = ExcelDataValidationOperator.greaterThan;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "greaterthanequal") {
                    validation.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "lessthan") {
                    validation.Operator = ExcelDataValidationOperator.lessThan;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "lessthanequal") {
                    validation.Operator = ExcelDataValidationOperator.lessThanOrEqual;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "notequal") {
                    validation.Operator = ExcelDataValidationOperator.notEqual;
                    validation.Formula.Value = ValA;
                } else {
                    throw new System.NullReferenceException("Data Validation Operator is not valid!");
                }

            return package.GetAsByteArray();
        }
    }

    public byte[] Data_Validation_Decimal(byte[] excelBinary, CellDataValidation cellDataValidation, DataValidation dataValidation) {
        using (var package = Excel_Open(excelBinary))
        {
            // Console.WriteLine(dataValidation.dataValidationConfig.ValidationOperator.ToLower());

            double ValA = 0;
            double ValB = 0;

            if(double.TryParse(dataValidation.FormulaValue1, out ValA)) {
                ValA = Convert.ToDouble(dataValidation.FormulaValue1);
            } else {
                throw new System.NullReferenceException("Formula 1 Value is not Decimal / Double!");
            }

            if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "between" || dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "not between") {
                if(double.TryParse(dataValidation.FormulaValue2, out ValB)) {
                    ValB = Convert.ToDouble(dataValidation.FormulaValue2);
                } else {
                    throw new System.NullReferenceException("Formula 2 Value is not Decimal / Double!");
                }

                // Console.WriteLine("Val A: " + ValA);
                // Console.WriteLine("Val B: " + ValB);

                if( ! (ValA >= 0 && ValB > 0 && ValB > ValA)) {
                    throw new System.NullReferenceException("Formula 1 and Formula 2 (Greater Than) Value required for between and not between Operator!");
                }

            }

            ExcelWorksheet worksheet = Worksheet_Select(package, cellDataValidation.SheetName);

            ExcelRange? targetExcelRange;
            try {
                targetExcelRange = Cell_Selections(package, cellDataValidation.CellRange.StartCellRow, cellDataValidation.CellRange.StartCellColumn, cellDataValidation.CellRange.EndCellRow, cellDataValidation.CellRange.EndCellColumn, cellDataValidation.CellName, cellDataValidation.SheetName);
            } catch {
                return package.GetAsByteArray();
            }

            if(targetExcelRange == null) {
                return package.GetAsByteArray();
            }


            worksheet.DataValidations.AddDecimalValidation(targetExcelRange.Address);
            ExcelDataValidationDecimal validation = (OfficeOpenXml.DataValidation.ExcelDataValidationDecimal) Data_Validation_Config(worksheet, targetExcelRange.Address, dataValidation.dataValidationConfig);

            // //Between (Default), Equal, GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, NotBetween, NotEqual

            if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "between" || dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "not between") { 
                validation.Operator = ExcelDataValidationOperator.between;
                validation.Formula.Value = ValA;
                validation.Formula2.Value = ValB;
            }  else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "equal") {
                    validation.Operator = ExcelDataValidationOperator.equal;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "greaterthan") {
                    validation.Operator = ExcelDataValidationOperator.greaterThan;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "greaterthanequal") {
                    validation.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "lessthan") {
                    validation.Operator = ExcelDataValidationOperator.lessThan;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "lessthanequal") {
                    validation.Operator = ExcelDataValidationOperator.lessThanOrEqual;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "notequal") {
                    validation.Operator = ExcelDataValidationOperator.notEqual;
                    validation.Formula.Value = ValA;
                } else {
                    throw new System.NullReferenceException("Data Validation Operator is not valid!");
                }

            return package.GetAsByteArray();
        }
    }

     public byte[] Data_Validation_DateTime(byte[] excelBinary, CellDataValidation cellDataValidation, DataValidation dataValidation) {
        using (var package = Excel_Open(excelBinary))
        {
            // Console.WriteLine(dataValidation.dataValidationConfig.ValidationOperator.ToLower());

            DateTime ValA = DateTime.Now;
            DateTime ValB = ValA.AddDays(1);

            if(DateTime.TryParse(dataValidation.FormulaValue1, out ValA)) {
                ValA = Convert.ToDateTime(dataValidation.FormulaValue1);
            } else {
                throw new System.NullReferenceException("Formula 1 Value is not Date Time!");
            }

            if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "between" || dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "not between") {
                if(DateTime.TryParse(dataValidation.FormulaValue2, out ValB)) {
                    ValB = Convert.ToDateTime(dataValidation.FormulaValue2);
                } else {
                    throw new System.NullReferenceException("Formula 2 Value is not Date Time!");
                }

                // Console.WriteLine("Val A: " + ValA);
                // Console.WriteLine("Val B: " + ValB);

                if( ! (ValB > ValA)) {
                    throw new System.NullReferenceException("Formula 1 and Formula 2 (Greater Than) Value required for between and not between Operator!");
                }

            }

            ExcelWorksheet worksheet = Worksheet_Select(package, cellDataValidation.SheetName);

            ExcelRange? targetExcelRange;
            try {
                targetExcelRange = Cell_Selections(package, cellDataValidation.CellRange.StartCellRow, cellDataValidation.CellRange.StartCellColumn, cellDataValidation.CellRange.EndCellRow, cellDataValidation.CellRange.EndCellColumn, cellDataValidation.CellName, cellDataValidation.SheetName);

            } catch {
                return package.GetAsByteArray();
            }

            if(targetExcelRange == null) {
                return package.GetAsByteArray();
            }


            worksheet.DataValidations.AddDateTimeValidation(targetExcelRange.Address);
            ExcelDataValidationDateTime validation = (OfficeOpenXml.DataValidation.ExcelDataValidationDateTime) Data_Validation_Config(worksheet, targetExcelRange.Address, dataValidation.dataValidationConfig);

            // //Between (Default), Equal, GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, NotBetween, NotEqual

            if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "between" || dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "not between") { 
                validation.Operator = ExcelDataValidationOperator.between;
                validation.Formula.Value = ValA;
                validation.Formula2.Value = ValB;
            }  else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "equal") {
                    validation.Operator = ExcelDataValidationOperator.equal;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "greaterthan") {
                    validation.Operator = ExcelDataValidationOperator.greaterThan;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "greaterthanequal") {
                    validation.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "lessthan") {
                    validation.Operator = ExcelDataValidationOperator.lessThan;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "lessthanequal") {
                    validation.Operator = ExcelDataValidationOperator.lessThanOrEqual;
                    validation.Formula.Value = ValA;
                } else if (dataValidation.dataValidationConfig.ValidationOperator.ToLower() == "notequal") {
                    validation.Operator = ExcelDataValidationOperator.notEqual;
                    validation.Formula.Value = ValA;
                } else {
                    throw new System.NullReferenceException("Data Validation Operator is not valid!");
                }

            return package.GetAsByteArray();
        }
    }   




    public byte[] Data_Validation_List(byte[] excelBinary, CellDataValidation cellDataValidation, DataValidationListItem dataValidationListItem) {

        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, cellDataValidation.SheetName);

            ExcelRange? targetExcelRange;
            try {
                targetExcelRange = Cell_Selections(package, cellDataValidation.CellRange.StartCellRow, cellDataValidation.CellRange.StartCellColumn, cellDataValidation.CellRange.EndCellRow, cellDataValidation.CellRange.EndCellColumn, cellDataValidation.CellName, cellDataValidation.SheetName);
            } catch {
                return package.GetAsByteArray();
            }
            
            if(targetExcelRange == null) {
                return package.GetAsByteArray();
            }

            worksheet.DataValidations.AddListValidation(targetExcelRange.Address);
            ExcelDataValidationList validation = (OfficeOpenXml.DataValidation.ExcelDataValidationList) Data_Validation_Config(worksheet, targetExcelRange.Address, dataValidationListItem.dataValidationConfig);

            if(dataValidationListItem.IsUsingItemFormula) {
                if (dataValidationListItem.ItemFormula != "") {
                    validation.Formula.ExcelFormula = dataValidationListItem.ItemFormula;                    
                    return package.GetAsByteArray();
                }
            } else {
                if(dataValidationListItem.ItemList != null) {
                    for (int i = 0; i < dataValidationListItem.ItemList.Length; i++)
                    {
                        validation.Formula.Values.Add(dataValidationListItem.ItemList[i]);
                    }
                    return package.GetAsByteArray();    
                } else {
                    throw new System.NullReferenceException("Item List is Empty!");
                }
            }

            return excelBinary;
        }
    }

// ============================================================
// Public Method Implementation Interface - Range Functions
// ============================================================

    public byte[] Range_Format(byte[] excelBinary, RangeFormat[] rangeFormats)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelRange? excelRanges;
            foreach(RangeFormat rangeFormat in rangeFormats) {
                try {
                    excelRanges = Cell_Selections(package, rangeFormat.CellRange.StartCellRow, rangeFormat.CellRange.StartCellColumn, rangeFormat.CellRange.EndCellRow, rangeFormat.CellRange.EndCellColumn, rangeFormat.CellName ?? "", rangeFormat.SheetName);

                } catch {
                    return package.GetAsByteArray();
                }
                
                if(excelRanges == null) {
                    return package.GetAsByteArray();
                }
                Cell_Format(excelRanges, rangeFormat.CellFormat);
            }
            return package.GetAsByteArray();
        }
    }

    public byte[] Range_BorderFormat(byte[] excelBinary, RangeBorderFormat[] rangeBorderFormats)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelRange? excelRanges;

            foreach(RangeBorderFormat rangeBorderFormat in rangeBorderFormats) {
                try {
                    excelRanges = Cell_Selections(package, rangeBorderFormat.CellRange.StartCellRow, rangeBorderFormat.CellRange.StartCellColumn, rangeBorderFormat.CellRange.EndCellRow, rangeBorderFormat.CellRange.EndCellColumn, rangeBorderFormat.CellName ?? "", rangeBorderFormat.SheetName);
                } catch {
                    return package.GetAsByteArray();
                }

                ExcelWorksheet worksheet = Worksheet_Select(package, rangeBorderFormat.SheetName);
                if(excelRanges == null) {
                    // Console.WriteLine("Range Null");
                    return package.GetAsByteArray();
                }
                worksheet = Border_Format(excelRanges, rangeBorderFormat.borderStyleFormat);
            }
            return package.GetAsByteArray();
        }
    }

    public RangeCellValue[] Range_CellRead(byte[] excelBinary, RangeCellRead[] rangeCellReads)
    {
        List<RangeCellValue> rangeCellValues = new List<RangeCellValue>();
         
        using (var package = Excel_Open(excelBinary))
        {
            ExcelRange? excelRanges;
            foreach(RangeCellRead rangeCellRead in rangeCellReads) {
                try {
                    excelRanges = Cell_Selections(package, rangeCellRead.CellRange.StartCellRow, rangeCellRead.CellRange.StartCellColumn, rangeCellRead.CellRange.EndCellRow, rangeCellRead.CellRange.EndCellColumn, rangeCellRead.CellName ?? "", rangeCellRead.SheetName);
                } catch {
                    return rangeCellValues.ToArray();
                }

                if(excelRanges != null) {
                    foreach(var row in excelRanges) {
                        foreach(var column in row) {
                            string? cellValueVar = Convert.ToString(column.Value);
                            if (cellValueVar != null) {
                                string[] cellStr = SplitRegex(row.Address);
                                int cellRowVar = Int32.Parse(cellStr[1]);
                                int cellColumnVar = columnNumber(cellStr[0]);
                                RangeCellValue rangeCellValue = new RangeCellValue();
                                rangeCellValue.CellRow = cellRowVar;
                                rangeCellValue.Value = cellValueVar;
                                rangeCellValue.CellColumn = cellColumnVar;
                                rangeCellValue.CellName = row.Address;
                                rangeCellValues.Add(rangeCellValue);
                            };
                        }
                    }
                }
            }
            return rangeCellValues.ToArray();
        }
    }


// ============================================================
// Public Method Implementation Interface - Miscellaneous
// ============================================================

    public byte[] Data_WriteJSON(byte[] excelBinary, DataWriteJSON[] dataWriteJSONs) {
        using (var package = Excel_Open(excelBinary)) {
            ExcelRange? excelRanges;

            foreach(DataWriteJSON dataWriteJSON in dataWriteJSONs) {
//                var jsonItems = System.Text.Json.JsonSerializer.Deserialize<IEnumerable<System.Dynamic.ExpandoObject>>(dataWriteJSON.JSONString);
                var jsonItems = JsonConvert.DeserializeObject<IEnumerable<System.Dynamic.ExpandoObject>>(dataWriteJSON.JSONString);

                try {
                    excelRanges = Cell_Selections(package, dataWriteJSON.Cell.CellRow, dataWriteJSON.Cell.CellColumn, 0, 0, dataWriteJSON.CellName ?? "", dataWriteJSON.SheetName);
                } catch {
                    return package.GetAsByteArray();
                }
                
                
                if(excelRanges == null) {
                    return package.GetAsByteArray();
                }

                ExcelRangeBase tableRange = excelRanges.LoadFromDictionaries(jsonItems, c =>
                { 
                    if(dataWriteJSON.IsShowHeader) {
                        c.PrintHeaders = true;
                    }

                    c.HeaderParsingType = HeaderParsingTypes.CamelCaseToSpace;

                    if(dataWriteJSON.TableStyle != "" || dataWriteJSON.TableStyle != null) {

                        TableStyles tableStyles;
                        if (Enum.TryParse<TableStyles>(dataWriteJSON.TableStyle, out tableStyles))
                        {
                            c.TableStyle = tableStyles;
                        }
                    }
                });
                if(dataWriteJSON.IsAutoFitColumn) {
                    tableRange.AutoFitColumns();
                }                
            }
            return package.GetAsByteArray();
        }
    }



    public byte[] Image_Insert(byte[] excelBinary, byte[] imageFile, int imageSizePercent = 100, int imageWidth = 0, int imageHeight = 0, int cellRow = 0, int cellColumn = 0, string? cellName = null, string? sheetName = null ) {
        if(imageFile.Length <= 0) return excelBinary;        
        using (var package = Excel_Open(excelBinary))
        {
            ExcelWorksheet worksheet = Worksheet_Select(package, sheetName);

            ExcelRange? excelRanges;
            try {
                excelRanges = Cell_Selections(package, cellRow, cellColumn, 0, 0, cellName ?? "", sheetName);
            } catch {
                return package.GetAsByteArray();
            }
            
            if(excelRanges == null) {
                return package.GetAsByteArray();
            }

            Stream stream = new MemoryStream(imageFile);
            ExcelPicture pic = worksheet.Drawings.AddPicture(GenerateGUID(), stream);

            string[] cellStr = SplitRegex(excelRanges.Address);

            int cellRowVar = Int32.Parse(cellStr[1]);
            int cellColumnVar = columnNumber(cellStr[0]);

            pic.SetPosition(cellRowVar, 0, cellColumnVar, 0);
            if(imageHeight > 0 && imageWidth > 0) {
                pic.SetSize(imageWidth, imageHeight);
            } else {
                pic.SetSize(imageSizePercent);  
            }
            
            return package.GetAsByteArray();
        }
    }

    public byte[] Chart_Create(byte[] excelBinary, string? sheetName)
    {
        throw new NotImplementedException();
    }

    public byte[] Excel_Merge(ExcelMerge[] ExcelFiles)
    {
        List<string> sheetNames = new List<string>();
        int dupCount = 0;
        using (ExcelPackage resultExcel = new ExcelPackage())
        {
            foreach(ExcelMerge excelFile in ExcelFiles) {
                using (var package = Excel_Open(excelFile.ExcelBinary)) {
                    for(int sheetIndex=0; sheetIndex<package.Workbook.Worksheets.Count; sheetIndex++) {
                        ExcelWorksheet excelWorksheet = package.Workbook.Worksheets[sheetIndex];
                        string sheetName = excelWorksheet.Name;
                        // Console.WriteLine("Sheet: " + sheetName);
                        foreach(string str in sheetNames) {
                            if(str == sheetName) {
                                sheetName = excelWorksheet.Name + dupCount.ToString();
                                dupCount++;
                                // Console.WriteLine("Sheet: " + sheetName + " Duplicate!");
                            };
                        }
                        sheetNames.Add(sheetName);
                        // Console.WriteLine(sheetNames.Count);
                        resultExcel.Workbook.Worksheets.Add(sheetName, excelWorksheet);
                    }
                }            
            }
            return resultExcel.GetAsByteArray();
        }
    }

// ============================================================
// Public Method Implementation Interface - Comment
// ============================================================

    public byte[] Comment_Add(byte[] excelBinary, CommentAdd[] commentAdds)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelRange? excelRanges;

            foreach(CommentAdd commentAdd in commentAdds) {
                ExcelWorksheet worksheet = Worksheet_Select(package, commentAdd.SheetName);

                try {
                    excelRanges = Cell_Selections(package, commentAdd.Cell.CellRow, commentAdd.Cell.CellColumn, 0, 0, commentAdd.CellName ?? "", commentAdd.SheetName);
                } catch {
                    return package.GetAsByteArray();
                }
                
                if(excelRanges == null) {
                    return package.GetAsByteArray();
                }

                excelRanges.AddComment(commentAdd.Text, commentAdd.Author);

            }
            return package.GetAsByteArray();
        }        
    }

    public byte[] Comment_Delete(byte[] excelBinary, CommentDelete[] commentDeletes)
    {
        using (var package = Excel_Open(excelBinary))
        {
            ExcelRange? excelRanges;

            foreach(CommentDelete commentDelete in commentDeletes) {
                ExcelWorksheet worksheet = Worksheet_Select(package, commentDelete.SheetName);

                try {
                    excelRanges = Cell_Selections(package, commentDelete.Cell.CellRow, commentDelete.Cell.CellColumn, 0, 0, commentDelete.CellName ?? "", commentDelete.SheetName);
                } catch {
                    return package.GetAsByteArray();
                }
                
                if(excelRanges == null) {
                    return package.GetAsByteArray();
                }

                worksheet.Comments.Remove(excelRanges.Comment);

            }
            return package.GetAsByteArray();
        }        
    }

}