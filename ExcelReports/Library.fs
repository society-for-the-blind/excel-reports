module ExcelReports.OIB

// #r "nuget: NPOI, 2.6.2"

open NPOI

type ICell             = SS.UserModel.ICell
type IRow              = SS.UserModel.IRow
type ISheet            = SS.UserModel.ISheet
type XSSFWorkbook      = XSSF.UserModel.XSSFWorkbook
type CellAddress       = SS.Util.CellAddress
type MissingCellPolicy = SS.UserModel.MissingCellPolicy
type CellRangeAddress  = SS.Util.CellRangeAddress
type CellReference     = SS.Util.CellReference


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

let updateCell (workbook: XSSFWorkbook) (sheetIndex: int) (rowIndex: int) (columnIndex: int) (value: string) =
    let sheet: ISheet = workbook.GetSheetAt(sheetIndex)
    let row: IRow = sheet.GetRow(rowIndex)
    let cell: ICell = row.GetCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK)
    cell.SetCellValue value

// TODO delete this at one point
let updateCellValue (cell: ICell) (value: string) =
    cell.SetCellValue value

let saveWorkbook (workbook: XSSFWorkbook) (filePath: string) =
    use fileStream: System.IO.FileStream =
        new System.IO.FileStream(
            filePath,
            System.IO.FileMode.Create, System.IO.FileAccess.Write
        )
    workbook.Write(fileStream)

// TODO move to a utility module
// ------------------------------
// TODO This needs to be a constrained type (and possibly renamed to
// `CellRecord` or smth) to be able to convert back and forth between
// NPOI's `CellAddress` and this type.
// -> NOTE 2023-11-19_19-56
//    Perhaps not. A constrained type would take a lot of time, and it now
//    conveys that it may fail to become an NPOI `CellAddress`. Converting
//    an NPOI `CellAddress` to a record will be simple, and it will be just
//    that, a temporary record.
// ------------------------------
type ProspectiveCellAddress = {
    Column: string;
    Row: int
}

// TODO move to a utility module
let maybeCellAddress ({ Column = col; Row = row}: ProspectiveCellAddress) :CellAddress option =
    let npoiCellAddress = (new CellAddress (col + (string row)))
    match npoiCellAddress with
    | _ as c when c.Column = -1 -> None
    | _ as c -> Some c

// TODO move to a utility module
let getCellAddressOrThrow : ProspectiveCellAddress -> CellAddress =
    Option.get << maybeCellAddress

let getCellByAddress (sheet: ISheet) (cellAddress: CellAddress) =
    let row: IRow = sheet.GetRow(cellAddress.Row)
    row.GetCell(cellAddress.Column, MissingCellPolicy.CREATE_NULL_AS_BLANK)

// let getCell (workbook: XSSFWorkbook) (sheetIndex: int) (rowIndex: int) (columnIndex: int) =
//     let sheet: ISheet = workbook.GetSheetAt(sheetIndex)
//     let row: IRow = sheet.GetRow(rowIndex)
//     row.GetCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK)

let getProspectiveCell (workbook: XSSFWorkbook) (sheetNumber: int) (prospectiveCellAddress: ProspectiveCellAddress) =
    let sheet: ISheet = workbook.GetSheetAt(sheetNumber)
    let cellAddress = getCellAddressOrThrow prospectiveCellAddress
    getCellByAddress sheet cellAddress

let getCell (workbook: XSSFWorkbook) (sheetNumber: int) (address: string) =
    let sheet: ISheet = workbook.GetSheetAt(sheetNumber)
    let cellAddress = new CellAddress (address)
    getCellByAddress sheet cellAddress

let findMergedRegionByCell (cell: ICell) =
    let cellRef = new CellReference (cell)
    cell.Sheet.MergedRegions
    |> Seq.tryFind (fun region -> region.IsInRange cellRef)

let findMergedRegionByStringAddress (workbook: XSSFWorkbook) (sheetNumber: int) (address: string) =
    let cell = getCell workbook sheetNumber address
    findMergedRegionByCell cell

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



open Npgsql.FSharp

let connectionString = "postgres://postgres:XntSrCoEEZtiacZrx2m7jR5htEoEfYyoKncfhNmnPrLqPzxXTU5nxM@192.168.64.4:5432/lynx"
let q = connectionString |> Sql.connect |> Sql.query "select * from lynx_sipnote where id = 27555;" |> Sql.execute (fun read -> read.text "note")

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



let convertStringToInt (s: string) =
    match System.Int32.TryParse(s) with
    | true, i -> Some i
    | _ -> None

let getSheetDataValidations (workbook: XSSFWorkbook) (sheetIndex: int) =
    let sheet: ISheet = workbook.GetSheetAt(sheetIndex)
    sheet.GetDataValidations()

let iterateThroughCellsInWorksheet (sheet: ISheet) (f: ICell -> unit) =
    let rows = sheet.GetRowEnumerator()
    while rows.MoveNext() do
        let row = (rows.Current :?> IRow).GetEnumerator()
        while row.MoveNext() do
            let cell = row.Current
            f cell

let lazyIterateThroughCellsInWorksheet (sheet: ISheet) (f: ICell -> unit) =
    let rows = sheet.GetRowEnumerator()
    let rowSeq =
        Seq.unfold
            (fun (enumerator: System.Collections.IEnumerator) ->
                if enumerator.MoveNext()
                then Some(enumerator.Current :?> IRow, enumerator)
                else None)
            rows
    rowSeq
    |> Seq.iter (fun row ->
        let cellEnumerator = row.GetEnumerator()
        let cellSeq =
            Seq.unfold
                (fun (enumerator: System.Collections.IEnumerator) ->
                    if enumerator.MoveNext()
                    then Some(enumerator.Current :?> ICell, enumerator)
                    else None)
                cellEnumerator
        cellSeq
        |> Seq.iter f)

//  (fun i -> let dv = (((s.GetDataValidations ()) |> Seq.toArray)[i]) in (dv.ValidationConstraint, dv.Regions.CellRangeAddresses)) 16;;
