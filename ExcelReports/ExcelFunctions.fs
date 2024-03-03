module ExcelReports.ExcelFunctions

// #r "nuget: DocumentFormat.OpenXml, 2.8.1"
(*
#r "nuget: NPOI, 2.6.2"
#load "ExcelReports/ExcelFunctions.fs";;
open ExcelReports.ExcelFunctions;;
*)

open NPOI
open OIBTypes

type ICell             = SS.UserModel.ICell
type IRow              = SS.UserModel.IRow
type ISheet            = SS.UserModel.ISheet
type XSSFWorkbook      = XSSF.UserModel.XSSFWorkbook
type CellAddress       = SS.Util.CellAddress
type MissingCellPolicy = SS.UserModel.MissingCellPolicy
type CellRangeAddress  = SS.Util.CellRangeAddress
type CellReference     = SS.Util.CellReference
type AreaReference     = SS.Util.AreaReference
type IDataValidation   = SS.UserModel.IDataValidation
type ValidationType    = SS.UserModel.ValidationType
type IndexedColors     = SS.UserModel.IndexedColors
type FillPattern       = SS.UserModel.FillPattern

type ColumnName = string
type ReportCell = Result<System.IFormattable, string>
type ReportColumn = ColumnName * ReportCell
type ReportRow = ReportColumn list
type ReportSheetData = ReportRow seq

/// <summary>
/// Opens an Excel file and returns the workbook.
/// </summary>
/// <param name="filePath">The path of the Excel file.</param>
/// <returns>The Excel workbook.</returns>
let openExcelFileWithNPOI (filePath: string) : XSSFWorkbook =
    use fileStream: System.IO.FileStream =
        new System.IO.FileStream(
            filePath,
            System.IO.FileMode.Open, System.IO.FileAccess.Read
        )
    let workbook = new XSSFWorkbook(fileStream)
    workbook

// let updateCell (workbook: XSSFWorkbook) (sheetIndex: int) (rowIndex: int) (columnIndex: int) (value: string) =
//     let sheet: ISheet = workbook.GetSheetAt(sheetIndex)
//     let row: IRow = sheet.GetRow(rowIndex)
//     let cell: ICell = row.GetCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK)
//     cell.SetCellValue value

// let rgb: byte array = [|255uy; 192uy; 150uy|]

// let o = openExcelFileWithNPOI "ExcelReports/20231208_protected_7-OB_Report-Data-Collection-Tool_V2.xlsx";;

open System.Reflection

let cloneCellStyle (cell: NPOI.XSSF.UserModel.XSSFCell)  =
    let original = cell.CellStyle
    // printfn "ORIGINAL: %A" original.FontIndex
    let workbook = cell.Sheet.Workbook
    let copy = workbook.CreateCellStyle()
    let properties = original.GetType().GetProperties(BindingFlags.Public ||| BindingFlags.Instance)
    for prop in properties do
        // printfn "MAIN: %s --- %A" prop.Name (prop.GetValue(original))
        if prop.CanRead && prop.CanWrite then
            // printfn "IF: %s --- %A" prop.Name (prop.GetValue(original))
            let value = prop.GetValue(original)
            prop.SetValue(copy, value)
    // `FontIndex` can only be set by the `SetFont` method.
    // The line below solved  my style mismatch issues, but
    // more complicated styles may  reveal other issues; to
    // track down the culprit, the `printfn` statements can
    // be used  to show which  properties cannot be  set in
    // the `for` loop above.
    copy.SetFont <| original.GetFont(workbook)
    copy

let hexStringToRGB (hexString: string) =
    let isHexColor (color: string) =
        "^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$"
        |> fun regexPattern -> new System.Text.RegularExpressions.Regex(regexPattern)
        |> fun regex -> regex.IsMatch(color)
    match (isHexColor hexString) with
    | false ->
        failwith "Invalid hex color string."
    | true ->
        let hexString = hexString.TrimStart('#')
        let r = System.Convert.ToByte(hexString.Substring(0, 2), 16)
        let g = System.Convert.ToByte(hexString.Substring(2, 2), 16)
        let b = System.Convert.ToByte(hexString.Substring(4, 2), 16)
        [|r; g; b|]

// TODO: refactor `changeCellColor` and `resetCellColor` as they are almost identical.
let changeCellColor (cell: ICell) (rgb: byte array) =
    let xssfCell: NPOI.XSSF.UserModel.XSSFCell = cell :?> NPOI.XSSF.UserModel.XSSFCell
    let newCellStyle: NPOI.SS.UserModel.ICellStyle = cloneCellStyle xssfCell
    // newCellStyle.CloneStyleFrom(xssfCell.CellStyle)
    match newCellStyle with
    // match xssfCell.CellStyle with
    | :? NPOI.XSSF.UserModel.XSSFCellStyle as xssfCellStyle ->
        let color = new NPOI.XSSF.UserModel.XSSFColor(rgb)
        xssfCellStyle.FillForegroundXSSFColor <- color
        xssfCellStyle.FillPattern <- NPOI.SS.UserModel.FillPattern.SolidForeground
    | _ -> failwith "'newCellStyle' cannot be cast to XSSFCellStyle"
    xssfCell.CellStyle <- newCellStyle

let resetCellColor (cell: ICell) =
    let xssfCell: NPOI.XSSF.UserModel.XSSFCell = cell :?> NPOI.XSSF.UserModel.XSSFCell
    let newCellStyle: NPOI.SS.UserModel.ICellStyle = cloneCellStyle xssfCell
    // newCellStyle.CloneStyleFrom(xssfCell.CellStyle)
    match newCellStyle with
    // match xssfCell.CellStyle with
    | :? NPOI.XSSF.UserModel.XSSFCellStyle as xssfCellStyle ->
        xssfCellStyle.FillPattern <- NPOI.SS.UserModel.FillPattern.NoFill
    | _ -> failwith "'newCellStyle' cannot be cast to XSSFCellStyle"
    xssfCell.CellStyle <- newCellStyle

let updateCell (cell: ICell) (string: string) =
    cell.SetCellValue string

let recalculateFormulas (workbook: XSSFWorkbook) =
    workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll()

let saveWorkbook (filePath: string) (workbook: XSSFWorkbook) =
    recalculateFormulas workbook
    use fileStream: System.IO.FileStream =
        new System.IO.FileStream(
            filePath,
            System.IO.FileMode.Create, System.IO.FileAccess.Write
        )
    workbook.Write(fileStream)

let getCellByAddress (sheet: ISheet) (cellAddress: CellAddress) =
    let row: IRow = sheet.GetRow(cellAddress.Row)
    row.GetCell(cellAddress.Column, MissingCellPolicy.CREATE_NULL_AS_BLANK)

let getCell (workbook: XSSFWorkbook) (sheetNumber: int) (address: string) =
    let sheet: ISheet = workbook.GetSheetAt(sheetNumber)
    let cellAddress = new CellAddress (address)
    getCellByAddress sheet cellAddress

let setCellFillColor (workbook: XSSFWorkbook) (cell: ICell) (color: IndexedColors) =
    // cell.CellStyle
    let c: NPOI.XSSF.UserModel.XSSFColor = new NPOI.XSSF.UserModel.XSSFColor(color)
    let style = workbook.CreateCellStyle()
    style.FillPattern <- FillPattern.SolidForeground
    style.FillForegroundColor <- color.Index
    cell.CellStyle <- style

// TODO clean up side effect galore (e.g., `resetCellColor` - or at least make them explicit...
let fillRow
    (workbook: XSSFWorkbook)
    (sheetNumber: int)
    (rowNumber: string)
    (newRow: ReportRow)
    (cellTransforms: (NPOI.SS.UserModel.ICell -> unit) list)
    =

    printfn "fillRow"
    let errorColor = hexStringToRGB "#ffc096"
    newRow
    |> Seq.iter (
        fun ((column, result): ReportColumn) ->
            printfn "column: %s" column
            // let rowNum = string(i + 7)
            let cell = getCell workbook sheetNumber (column + rowNumber)

            // Some cells in a template require special care; see `generateQuarterlyReport` in `ExcelReports.fs`.
            List.iter (fun f -> f cell) cellTransforms

            let cellString =
                match result with
                | Ok reportValue ->
                    string reportValue
                | Error str ->
                    changeCellColor cell errorColor
                    "Error: " + str
            updateCell cell cellString
    )

let populateSheet
    (rows: ReportSheetData)
    (sheetNumber: int)
    (rowOffset: int)
    (cellTransforms: (NPOI.SS.UserModel.ICell -> unit) list)
    (workbook: XSSFWorkbook)
    : XSSFWorkbook
    =

    printfn "populateSheet"
    rows
    // |> Seq.sortBy extractClientName
    |> Seq.iteri (
        fun i row ->
             printfn "populateSheet iteri"
             // TODO "sheet_start_row"
             //      The  numbe r 7  below  denotes  the  start  row from
             //      where  the  "rows"  is  `ReportSheetData`  need  to  be
             //      pasted; both  the "demographics" and  the "services"
             //      sheets start from row 7, but it should probably be a
             //      parameter.
             fillRow workbook sheetNumber (string(i + rowOffset)) row cellTransforms
            //  fillRow workbook sheetNumber (string(i + 7)) row cellTransforms
       )
    workbook
