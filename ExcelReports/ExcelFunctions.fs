module ExcelReports.ExcelFunctions

// #r "nuget: DocumentFormat.OpenXml, 2.8.1"
(*
#r "nuget: NPOI, 2.6.2"
#load "ExcelReports/ExcelFunctions.fs";;
open ExcelReports.ExcelFunctions;;
*)

open NPOI

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

// NOTE 2023-11-27_10-24
// Shelving this experiment to create a more meaningful type for cell
// address strings, but not sure if it's worth to create an elaborate
// constrained type just yet.
//
// // TODO move to a utility module
// // ------------------------------
// // TODO This needs to be a constrained type (and possibly renamed to
// // `CellRecord` or smth) to be able to convert back and forth between
// // NPOI's `CellAddress` and this type.
// // -> NOTE 2023-11-19_19-56
// //    Perhaps not. A constrained type would take a lot of time, and it now
// //    conveys that it may fail to become an NPOI `CellAddress`. Converting
// //    an NPOI `CellAddress` to a record will be simple, and it will be just
// //    that, a temporary record.
// // ------------------------------
// type ProspectiveCellAddress = {
//     Column: string;
//     Row: int
// }
//
// // // TODO move to a utility module
// // let maybeCellAddress ({ Column = col; Row = row}: ProspectiveCellAddress) :CellAddress option =
// //     let npoiCellAddress = (new CellAddress (col + (string row)))
// //     match npoiCellAddress with
// //     | _ as c when c.Column = -1 -> None
// //     | _ as c -> Some c
//
// // // TODO move to a utility module
// // let getCellAddressOrThrow : ProspectiveCellAddress -> CellAddress =
// //     Option.get << maybeCellAddress
//
// let getProspectiveCell (workbook: XSSFWorkbook) (sheetNumber: int) (prospectiveCellAddress: ProspectiveCellAddress) =
//     let sheet: ISheet = workbook.GetSheetAt(sheetNumber)
//     let cellAddress = getCellAddressOrThrow prospectiveCellAddress
//     getCellByAddress sheet cellAddress

// NOTE Cells -> Row -> Sheet -> Workbook
let findMergedRegion (cell: ICell) =
    let cellRef = new CellReference (cell)
    cell.Sheet.MergedRegions
    |> Seq.tryFind (fun (region: CellRangeAddress) -> region.IsInRange cellRef)

let findMergedRegionByStringAddress (workbook: XSSFWorkbook) (sheetNumber: int) (address: string) =
    // (findMergedRegion << (getCell workbook sheetNumber)) address
    let cell: ICell = getCell workbook sheetNumber address
    findMergedRegion cell

let findDataValidation (cell: ICell) =
    let cellRef = new CellReference (cell)
    cell.Sheet.GetDataValidations()
    |> Seq.tryFind (fun (validation: SS.UserModel.IDataValidation) -> validation.Regions.CellRangeAddresses |> Seq.exists (fun region -> region.IsInRange cellRef))

let findDataValidationByCellAddress (workbook: XSSFWorkbook) (sheetNumber: int) (address: string) : SS.UserModel.IDataValidation option =
    // (findDataValidation << (getCell workbook sheetNumber)) address
    let cell: ICell = getCell workbook sheetNumber address
    findDataValidation cell

// Write a function to enumerate available values from an `NPOI.XSSF.UserModel.XSSFDataValidationConstraint` when the values are defined by a range of cells (e.g., "$B$2:$C$2")

let convertCellRangeToList (cellRange: string) =
    let arr = cellRange.Split(':')
    let startCellAddress = new CellAddress(arr.[0])
    let endCellAddress = new CellAddress(arr.[1])
    let startRow = startCellAddress.Row
    let endRow = endCellAddress.Row
    let startColumn = startCellAddress.Column
    let endColumn = endCellAddress.Column
    let cellAddresses =
        [ for row in startRow .. endRow do
            for column in startColumn .. endColumn do
                let cellAddress = new CellAddress(row, column)
                cellAddress.ToString() ]
    cellAddresses


let enumerateValuesFromRange (cell: ICell) =
    let validation = Option.get (findDataValidation cell)
    let vc = validation.ValidationConstraint
    let rangeAddress = vc.Formula1
    let arr = rangeAddress.Split('!')
    let sheet = cell.Sheet.Workbook.GetSheet(arr[0].TrimStart('='))
    let cellAddresses = arr[1].Split(':')
    let startCellAddress = new CellAddress(cellAddresses.[0])
    let endCellAddress = new CellAddress(cellAddresses.[1])
    let startRow = startCellAddress.Row
    let endRow = endCellAddress.Row
    let startColumn = startCellAddress.Column
    let endColumn = endCellAddress.Column
    let values =
        [ for row in startRow .. endRow do
            for column in startColumn .. endColumn do
                let cell = sheet.GetRow(row).GetCell(column)
                match cell with
                | null -> None
                | _ -> Some(cell.StringCellValue) ]
        |> List.choose id
    values


let t (cell: ICell) = cell.Sheet.Workbook
// workbook.GetAllNames() = Status, Pending, Assessed named cell ranges
// PONDER
// How to match the found MergedRegion for a cell with the data
// validations "defined" ("applicable"?) for that cell?

// let lofa (a: SS.Util.CellReference) = a.Row

// let cellRangeAddresstoRowTuple (cellRangeAddress: CellRangeAddress) : CellAddress * CellAddress =
//     ( (new CellAddress (cellRangeAddress.FirstRow, cellRangeAddress.FirstColumn)),
//       (new CellAddress (cellRangeAddress.LastRow, cellRangeAddress.LastColumn))
//     )

// let isCellAddressInRange (cellRangeAddress: CellRangeAddress) (cellAddress: CellAddress) =
//     if cellRangeAddress.FirstRow <> cellRangeAddress.LastRow then
//         failwith "NOT IMPLEMENTED: Cell range address spans multiple rows."
//     let (firstCellAddress, lastCellAddress) = cellRangeAddresstoCellAddressTuple cellRangeAddress
//     let firstCellRow = firstCellAddress.Row
//     let firstCellColumn = firstCellAddress.Column
//     let lastCellRow = lastCellAddress.Row
//     let lastCellColumn = lastCellAddress.Column
//     let cellRow = cellAddress.Row
//     let cellColumn = cellAddress.Column
//     (firstCellRow <= cellRow && cellRow <= lastCellRow) &&
//     (firstCellColumn <= cellColumn && cellColumn <= lastCellColumn)

let getSheetDataValidations (workbook: XSSFWorkbook) (sheetIndex: int) =
    let sheet: ISheet = workbook.GetSheetAt(sheetIndex)
    sheet.GetDataValidations()

//  (fun i -> let dv = (((s.GetDataValidations ()) |> Seq.toArray)[i]) in (dv.ValidationConstraint, dv.Regions.CellRangeAddresses)) 16;;

// VALIDATIONS

let getNamedRangeValues (workbook: XSSFWorkbook) (namedRangeName: string) =
    let namedRange = workbook.GetName(namedRangeName)
    let area = new AreaReference(namedRange.RefersToFormula, workbook.SpreadsheetVersion)
    // NOTE 2023-11-22_0740
    // These `cellRefs` retain reference to the sheet as well because
    // `RefersToFormula` is essentially a workbook level property.
    // cf. NOTE 2023-11-22_0743
    let cellRefs = area.GetAllReferencedCells()
    // let cellValues =
    cellRefs
    |> Array.map (fun cellRef ->
        let sheet = workbook.GetSheet(cellRef.SheetName)
        let cell = sheet.GetRow(cellRef.Row).GetCell((int cellRef.Col))
        cell.StringCellValue
    )
    // Array.contains newValue cellValues

let listWorksheets (workbook: XSSFWorkbook) =
    [|for i in 0 .. workbook.NumberOfSheets - 1 -> workbook.GetSheetAt(i)|]

let getAllDataValidations (workbook: XSSFWorkbook) =
    listWorksheets workbook
    |> Array.map (fun sheet ->
        sheet.GetDataValidations()
        |> Seq.toArray
    )
    |> Array.concat

// NOTE 2023-11-24_0712
// It was a conscious decision to focus only on list-type data
// validations, and throw if there are more.
// WHY? To curb complexity: https://stackoverflow.com/a/77518656/1498178
// The gist is that there are at least 7 types of data validations, but
// OIB reports only(?) use those - and list-type data validations are
// complex enough themselves.
let getDataValidationsForCell (cell: ICell)=
    let sheet = cell.Sheet
    sheet.GetDataValidations()
    |> Seq.filter (fun dv ->
        dv.Regions.CellRangeAddresses
        |> Seq.exists (fun cellRangeAddress -> cellRangeAddress.IsInRange(new CellReference(cell)))
    )
    |> Seq.toArray

// let getDataValidationValues (sheet: ISheet) (dv: IDataValidation) =

//====================

let a (s: SS.UserModel.IRichTextString) = 27
let b (s: SS.Util.AreaReference) = 27
let c (s: SS.Util.CellReference) = 27
let d (s: XSSFWorkbook) = s.GetAllNames()
let e (r: IRow) = 27

// type PasteDirection = | Down | Right // | Up | Left

// let insertListDown (startCell: ICell) (list: string list): unit =
// NOTE 2023-11-19_20-32 How to get the next cell in a column
// > let a7 = getCellByAddress (o.GetSheetAt 3) (getCellAddressOrThrow { Column = "A"; Row = 7});;
// val a7: ICell = Sample Joe Smith

// > let a8address = new CellAddress ( (a7.Address.Row + 1), a7.Address.Column);;
// val a8address: CellAddress = A8

// > let a8 = getCellByAddress a7.Sheet a8address;;
// val a8: ICell = Sample Ann Jones

// let insertListRight (sheet: ISheet) (startCellAddress: CellAddress) (list: string list): unit =
//     let row: IRow = sheet.GetRow(rowNumber)
//     let cell: ICell = row.GetCell(columnNumber, MissingCellPolicy.CREATE_NULL_AS_BLANK)
//     let cellAddress: CellAddress = new CellAddress(columnName)
//     let rowNumber = cellAddress.Row
//     let columnNumber = cellAddress.Column
//     let row: IRow = sheet.GetRow(rowNumber)
//     let cell: ICell = row.GetCell(columnNumber, MissingCellPolicy.CREATE_NULL_AS_BLANK)
//     let cellStyle = cell.CellStyle
//     let dataValidationHelper = sheet.GetDataValidationHelper()
//     let dataValidationConstraint = dataValidationHelper.CreateExplicitListConstraint(clients |> List.toArray)
//     let dataValidation = dataValidationHelper.CreateValidation(dataValidationConstraint, cellAddress)
//     dataValidation.ShowErrorBox <- true
//     dataValidation.ErrorStyle <- DataValidation.ErrorStyle.STOP
//     dataValidation.CreateErrorBox("Error", "Please select a valid client from the dropdown list.")
//     sheet.AddValidationData(dataValidation)
//     cell.CellStyle <- cellStyle

// let insertList (sheet: ISheet) (startCellAddress: CellAddress) (direction: PasteDirection) (list: string list): unit =
//     match direction with
//     | Down -> insertListDown sheet startCellAddress list
//     | Right -> insertListRight sheet startCellAddress list

// let enterClients (sheet: ISheet) (startCellAddress: CellAddress) (direction: PasteDirection) =

let setNamedRange (workbook: XSSFWorkbook) (sheetName: string) (rangeName: string) (startRow: int) (startColumn: int) (endRow: int) (endColumn: int) =
    let name: SS.UserModel.IName = workbook.CreateName()
    name.NameName <- rangeName
    let cellRangeAddress = new CellRangeAddress(startRow, endRow, startColumn, endColumn)
    name.RefersToFormula <- $"{sheetName}!{cellRangeAddress.FormatAsString()}"

// let writeRow values =