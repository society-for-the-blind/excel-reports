module ExcelReports.OIB

// #r "nuget: NPOI, 2.6.2"

open NPOI

type ICell             = SS.UserModel.ICell
type IRow              = SS.UserModel.IRow
type ISheet            = SS.UserModel.ISheet
type XSSFWorkbook      = XSSF.UserModel.XSSFWorkbook
type CellAddress       = SS.Util.CellAddress
type MissingCellPolicy = SS.UserModel.MissingCellPolicy

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

let getCell (workbook: XSSFWorkbook) (sheetIndex: int) (rowIndex: int) (columnIndex: int) =
    let sheet: ISheet = workbook.GetSheetAt(sheetIndex)
    let row: IRow = sheet.GetRow(rowIndex)
    row.GetCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK)

let saveWorkbook (workbook: XSSFWorkbook) (filePath: string) =
    use fileStream: System.IO.FileStream =
        new System.IO.FileStream(
            filePath,
            System.IO.FileMode.Create, System.IO.FileAccess.Write
        )
    workbook.Write(fileStream)

// TODO move to a utility module
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

type PasteDirection = | Down | Right // | Up | Left

let enterClients (sheet: ISheet) (startCellAddress: CellAddress) (direction: PasteDirection) =
    let columnNumber = column.ColumnIndex
    let rowNumber = column.RowIndex
    let row: IRow = sheet.GetRow(rowNumber)
    let cell: ICell = row.GetCell(columnNumber, MissingCellPolicy.CREATE_NULL_AS_BLANK)
    let cellAddress: CellAddress = new CellAddress(columnName)
    let rowNumber = cellAddress.Row
    let columnNumber = cellAddress.Column
    let row: IRow = sheet.GetRow(rowNumber)
    let cell: ICell = row.GetCell(columnNumber, MissingCellPolicy.CREATE_NULL_AS_BLANK)
    let cellStyle = cell.CellStyle
    let dataValidationHelper = sheet.GetDataValidationHelper()
    let dataValidationConstraint = dataValidationHelper.CreateExplicitListConstraint(clients |> List.toArray)
    let dataValidation = dataValidationHelper.CreateValidation(dataValidationConstraint, cellAddress)
    dataValidation.ShowErrorBox <- true
    dataValidation.ErrorStyle <- DataValidation.ErrorStyle.STOP
    dataValidation.CreateErrorBox("Error", "Please select a valid client from the dropdown list.")
    sheet.AddValidationData(dataValidation)
    cell.CellStyle <- cellStyle








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
