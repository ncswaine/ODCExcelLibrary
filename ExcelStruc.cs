using OutSystems.ExternalLibraries.SDK;

namespace OutSystems.ExternalLib.Excel
{


    [OSStructure]
    public struct TabProperties
    {
        public int ColorR;
        public int ColorG;
        public int ColorB;
        public int ColorA;
        public string ColorHex;
    }

    [OSStructure]
    public struct DimensionProperties
    {
        public string Address;
        public string AddressStart;
        public string AddressEnd;
        public string FullAddress;
        public int ColumnStart;
        public int ColumnEnd;
        public int RowStart;
        public int RowEnd;
    }

    [OSStructure]
    public struct WorksheetProperties
    {
        public string Name;
        public int Index;
        public TabProperties TabProperties;
        public DimensionProperties DimensionProperties;
    }

    [OSStructure]
    public struct Worksheet
    {
        public int Index { get; set; }
        public string Name { get; set; }

        string? colorHex;
        public string ColorHex { get { return colorHex ?? ""; } set { colorHex = value; } }
    }

    [OSStructure]
    public struct BorderStyleFormat
    {
        //DashDot, DashDotDot, Dashed, Dotted, Double, Hair, Medium, MediumDashDot, MediumDashDotDot, MediumDashed, Thick, Thin, None
        string? borderStyle;
        string? borderColorHex;
        bool? isTop;
        bool? isBottom;
        bool? isLeft;
        bool? isRight;
        bool? isRound;

        [OSStructureField(Description = "DashDot, DashDotDot, Dashed, Dotted, Double, Hair, Medium, MediumDashDot, MediumDashDotDot, MediumDashed, Thick, Thin, None", OriginalName = "BorderStyle")]
        public string BorderStyle { get { return borderStyle ?? "None"; } set { borderStyle = value; } }
        public string BorderColorHex { get { return borderColorHex ?? "#000000"; } set { borderColorHex = value; } }
        public bool IsTop { get { return isTop ?? false; } set { isTop = value; } }
        public bool IsBottom { get { return isBottom ?? false; } set { isBottom = value; } }
        public bool IsLeft { get { return isLeft ?? false; } set { isLeft = value; } }
        public bool IsRight { get { return isRight ?? false; } set { isRight = value; } }
        public bool IsRound { get { return isRound ?? false; } set { isRound = value; } }
    }


    [OSStructure]
    public struct FontStyleFormat
    {
        string? fortNameVar;
        int? fontSizeVar;
        string? fontColorHexVar;
        bool? isBoldVar;
        bool? isItalicVar;
        bool? isUnderlineVar;
        bool? isShrinkToFit;
        bool? isWrapText;
        bool? isQuotePrefix;
        string? horizontalAlignment;
        string? verticalAlignment;

        public string FontName { get { return fortNameVar ?? ""; } set { fortNameVar = value; } }
        public int FontSize { get { return fontSizeVar ?? 0; } set { fontSizeVar = value; } }
        public string FontColorHex { get { return fontColorHexVar ?? ""; } set { fontColorHexVar = value; } }
        public bool IsBold { get { return isBoldVar ?? false; } set { isBoldVar = value; } }
        public bool IsItalic { get { return isItalicVar ?? false; } set { isItalicVar = value; } }
        public bool IsUnderline { get { return isUnderlineVar ?? false; } set { isUnderlineVar = value; } }
        public bool IsShrinkToFit { get { return isShrinkToFit ?? false; } set { isShrinkToFit = value; } }
        public bool IsWrapText { get { return isWrapText ?? false; } set { isWrapText = value; } }
        public bool IsQuotePrefix { get { return isQuotePrefix ?? false; } set { isQuotePrefix = value; } }

        [OSStructureField(Description = "Left, Center, Right", OriginalName = "HorizontalAlignment")]
        public string HorizontalAlignment { get { return horizontalAlignment ?? ""; } set { horizontalAlignment = value; } }

        [OSStructureField(Description = "Top, Center, Bottom", OriginalName = "VerticalAlignment")]
        public string VerticalAlignment { get { return verticalAlignment ?? ""; } set { verticalAlignment = value; } }
    }


    [OSStructure]
    public struct CellFormat
    {

        string? cellTypeVar;
        string? cellTypeFormatVar;

        public FontStyleFormat FontStyleFormat { get; set; }
        string? backgroundColorHexVar;


        public bool IsLocked { get; set; }
        public bool IsHidden { get; set; }
        public bool IsAutoFitColumn { get; set; }


        [OSStructureField(Description = "Text (default), Number, Decimal, DateTime, Bool, Formula", OriginalName = "CellType")]
        public string CellType { get { return cellTypeVar ?? "Text"; } set { cellTypeVar = value; } }
        public string CellTypeFormat
        {
            get
            {
                if (CellType.ToLower() == "datetime" && cellTypeFormatVar == "")
                {
                    return "dd/mm/yyyy";
                }
                return cellTypeFormatVar ?? "@";
            }
            set { cellTypeFormatVar = value; }
        }

        public string BackgroundColorHex { get { return backgroundColorHexVar ?? ""; } set { backgroundColorHexVar = value; } }
    }

    [OSStructure]
    public struct Cell
    {
        int? cellRow;
        int? cellColumn;

        public int CellRow { get { return cellRow ?? 0; } set { cellRow = value; } }
        public int CellColumn { get { return cellColumn ?? 0; } set { cellColumn = value; } }
    }

    [OSStructure]
    public struct CellRange
    {
        int? startCellRow;
        int? startCellColumn;
        int? endCellRow;
        int? endCellColumn;

        public int StartCellRow { get { return startCellRow ?? 0; } set { startCellRow = value; } }
        public int StartCellColumn { get { return startCellColumn ?? 0; } set { startCellColumn = value; } }
        public int EndCellRow { get { return endCellRow ?? 0; } set { endCellRow = value; } }
        public int EndCellColumn { get { return endCellColumn ?? 0; } set { endCellColumn = value; } }

    }

    [OSStructure]
    public struct CellWrite
    {
        [OSStructureField(IsMandatory = true, OriginalName = "CellValue", Description = "For Formula, Don't use semicolon as a separator between function arguments. Only comma is supported.")]
        public string CellValue { get; set; }
        public Cell Cell { get; set; }
        string? cellName;
        string? sheetName;
        public CellFormat CellFormat { get; set; }

        public string CellName { get { return cellName ?? ""; } set { cellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }
    }

    [OSStructure]
    public struct CellMerge
    {
        public CellRange CellRange { get; set; }
        string? cellName;
        string? sheetName;

        public string CellName { get { return cellName ?? ""; } set { cellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }
    }


    [OSStructure]
    public struct RangeFormat
    {
        public CellFormat CellFormat { get; set; }
        public CellRange CellRange { get; set; }
        string? cellName;
        string? sheetName;
        public string CellName { get { return cellName ?? ""; } set { cellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }
    }

    [OSStructure]
    public struct DataWriteJSON
    {
        [OSStructureField(IsMandatory = true, OriginalName = "JSONString")]
        public string JSONString { get; set; }
        bool? isShowHeader;
        public Cell Cell { get; set; }
        string? cellName;
        string? sheetName;
        string? tableStyle;
        public bool IsAutoFitColumn { get; set; }


        public bool IsShowHeader { get { return isShowHeader ?? false; } set { isShowHeader = value; } }

        public string CellName { get { return cellName ?? ""; } set { cellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }

        [OSStructureField(Description = "Based on https://epplussoftware.com/docs/5.2/api/OfficeOpenXml.Table.TableStyles.html", OriginalName = "TableStyle")]
        public string TableStyle { get { return tableStyle ?? ""; } set { tableStyle = value; } }
    }

    [OSStructure]
    public struct RangeBorderFormat
    {
        public BorderStyleFormat borderStyleFormat { get; set; }
        public CellRange CellRange { get; set; }
        string? cellName;
        string? sheetName;
        public string CellName { get { return cellName ?? ""; } set { cellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }
    }

    [OSStructure]
    public struct CellCopy
    {
        public Cell SourceCell { get; set; }
        string? sourceCellName;
        public CellRange DestinationCell { get; set; }
        string? destinationCellName;
        string? sheetName;

        public string SourceCellName { get { return sourceCellName ?? ""; } set { sourceCellName = value; } }
        public string DestinationCellName { get { return destinationCellName ?? ""; } set { destinationCellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }

        public readonly bool IsEmpty()
        {
            if (((SourceCell.CellRow == 0 || SourceCell.CellColumn == 0) && (sourceCellName == "" || sourceCellName == null)) ||
            ((DestinationCell.StartCellRow == 0 || DestinationCell.StartCellColumn == 0 || DestinationCell.EndCellRow == 0 || DestinationCell.EndCellRow == 0) && (destinationCellName == "" || destinationCellName == null)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }

    [OSStructure]
    public struct CellFindResult
    {
        public Cell Cell { get; set; }
        public string CellName { get; set; }
        public string CellValue { get; set; }
    }

    [OSStructure]
    public struct CommentAdd
    {
        public Cell Cell { get; set; }
        string? cellName;

        [OSStructureField(IsMandatory = true, OriginalName = "Text")]
        public string Text { get; set; }

        [OSStructureField(IsMandatory = true, OriginalName = "Author")]
        public string Author { get; set; }
        string? sheetName;

        public string CellName { get { return cellName ?? ""; } set { cellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }
    }

    [OSStructure]
    public struct CommentDelete
    {
        public Cell Cell { get; set; }
        string? cellName;
        string? sheetName;

        public string CellName { get { return cellName ?? ""; } set { cellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }
    }

    [OSStructure]
    public struct KeyValue
    {
        [OSStructureField(IsMandatory = true, OriginalName = "Key")]
        public string Key { get; set; }
        [OSStructureField(IsMandatory = true, OriginalName = "Value")]
        public string Value { get; set; }
    }

    [OSStructure]
    public struct WorkbookProperties
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public string Comments { get; set; }
        public string Company { get; set; }
        public string Subject { get; set; }
        public string Manager { get; set; }
        public string Category { get; set; }
        public string Keywords { get; set; }

        public KeyValue[] KeyValues { get; set; }
    }

    [OSStructure]
    public struct CellDataValidation
    {
        public CellRange CellRange { get; set; }
        string? cellName;
        string? sheetName;
        public string CellName { get { return cellName ?? ""; } set { cellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }
    }


    [OSStructure]
    public struct DataValidationConfig
    {
        string? validationOperator;
        bool? isShowErrorMessage;
        string? errorStyle;
        string? errorTitle;
        string? errorMessage;
        bool? isShowInputMessage;
        string? inputTitle;
        string? inputMessage;

        [OSStructureField(Description = "Equal (Default), GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, Between, NotBetween, NotEqual", OriginalName = "ValidationOperator")]
        public string ValidationOperator { get { return validationOperator ?? "Equal"; } set { validationOperator = value; } }

        public bool IsShowErrorMessage { get { return isShowErrorMessage ?? true; } set { isShowErrorMessage = value; } }

        [OSStructureField(Description = "Information, Stop, Warning (default)", OriginalName = "ErrorStyle")]
        public string ErrorStyle { get { return errorStyle ?? "Warning"; } set { errorStyle = value; } }

        public string ErrorTitle { get { return errorTitle ?? ""; } set { errorTitle = value; } }
        public string ErrorMessage { get { return errorMessage ?? ""; } set { errorMessage = value; } }
        public bool IsShowInputMessage { get { return isShowInputMessage ?? false; } set { isShowInputMessage = value; } }
        public string InputTitle { get { return inputTitle ?? ""; } set { inputTitle = value; } }
        public string InputMessage { get { return inputMessage ?? ""; } set { inputMessage = value; } }
    }


    [OSStructure]
    public struct DataValidationListItem
    {

        public string[]? ItemList;
        string? itemFormula;
        bool? isUsingItemFormula;
        public DataValidationConfig dataValidationConfig;

        [OSStructureField(Description = "Use Cell Lock! Example: $B$1:#B$10 or 'Sheet 1'!$B$6:$B$10", OriginalName = "ItemFormula")]
        public string ItemFormula { get { return itemFormula ?? ""; } set { itemFormula = value; } }

        [OSStructureField(Description = "Default value is FALSE, Set TRUE, if you want to use ItemFormula as List", OriginalName = "IsUsingItemFormula")]
        public bool IsUsingItemFormula { get { return isUsingItemFormula ?? false; } set { isUsingItemFormula = value; } }
    }


    [OSStructure]
    public struct DataValidation
    {
        [OSStructureField(IsMandatory = true, OriginalName = "FormulaValue1")]
        public string FormulaValue1;
        string? formulaValue2;

        public DataValidationConfig dataValidationConfig;

        [OSStructureField(Description = "Only Applicable for between and not berween Operator! Must be Greater than Formula 1 Value!", OriginalName = "FormulaValue2")]
        public string FormulaValue2 { get { return formulaValue2 ?? ""; } set { formulaValue2 = value; } }
    }


    [OSStructure]
    public struct RichTextFormatText
    {
        [OSStructureField(IsMandatory = true, OriginalName = "CellValue")]
        public string CellValue { get; set; }

        string? fortNameVar;
        int? fontSizeVar;
        string? fontColorHexVar;
        bool? isBoldVar;
        bool? isItalicVar;
        bool? isUnderlineVar;
        bool? isStrikeOutVar;

        [OSStructureField(Description = "Default Font Name is Calibri", OriginalName = "FontName")]
        public string FontName { get { return fortNameVar ?? ""; } set { fortNameVar = value; } }

        [OSStructureField(Description = "Default Font Size is 11", OriginalName = "FontSize")]
        public int FontSize { get { return fontSizeVar ?? 0; } set { fontSizeVar = value; } }

        [OSStructureField(Description = "Default Font Color is Black", OriginalName = "FontColorHex")]
        public string FontColorHex { get { return fontColorHexVar ?? ""; } set { fontColorHexVar = value; } }


        public bool IsBold { get { return isBoldVar ?? false; } set { isBoldVar = value; } }
        public bool IsItalic { get { return isItalicVar ?? false; } set { isItalicVar = value; } }
        public bool IsUnderline { get { return isUnderlineVar ?? false; } set { isUnderlineVar = value; } }
        public bool IsStrikeOut { get { return isStrikeOutVar ?? false; } set { isStrikeOutVar = value; } }
    }


    [OSStructure]
    public struct CellWriteRichText
    {
        public Cell Cell { get; set; }
        string? cellName;
        string? sheetName;
        public bool IsAutoFitColumn { get; set; }
        string? horizontalAlignment;
        string? verticalAlignment;

        public RichTextFormatText[] RichTextFormatTexts { get; set; }

        public string CellName { get { return cellName ?? ""; } set { cellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }

        [OSStructureField(Description = "Left, Center, Right", OriginalName = "HorizontalAlignment")]
        public string HorizontalAlignment { get { return horizontalAlignment ?? ""; } set { horizontalAlignment = value; } }

        [OSStructureField(Description = "Top, Center, Bottom", OriginalName = "VerticalAlignment")]
        public string VerticalAlignment { get { return verticalAlignment ?? ""; } set { verticalAlignment = value; } }


    }

    [OSStructure]
    public struct RangeCellValue
    {
        int? cellRow;
        int? cellColumn;

        public string Value;
        public string CellName;

        public int CellRow { get { return cellRow ?? 0; } set { cellRow = value; } }
        public int CellColumn { get { return cellColumn ?? 0; } set { cellColumn = value; } }
    }

    [OSStructure]
    public struct RangeCellRead
    {
        public CellRange CellRange { get; set; }
        string? cellName;
        string? sheetName;
        public string CellName { get { return cellName ?? ""; } set { cellName = value; } }
        public string SheetName { get { return sheetName ?? ""; } set { sheetName = value; } }
    }

    [OSStructure]
    public struct ExcelMerge
    {
        public byte[] ExcelBinary { get; set; }
    }

    [OSStructure]
    public struct ExcelImages
    {
        public byte[] Image { get; set; }
        public string ImageName { get; set; }
        public Cell Cell { get; set; }
        public string CellName { get; set; }
        public string SheetName { get; set; }
        string? imageType { get; set; }
        public string ImageType { get { return imageType ?? ""; } set { imageType = value; } }
        public double ImageWidth { get; set; }
        public double ImageHeight { get; set; }
    }


}