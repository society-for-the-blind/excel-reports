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
type AreaReference     = SS.Util.AreaReference
type IDataValidation   = SS.UserModel.IDataValidation
type ValidationType    = SS.UserModel.ValidationType

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

let updateCell (cell: ICell) (string: string) =
    cell.SetCellValue string

let recalculateFormulas (workbook: XSSFWorkbook) =
    workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll()

let saveWorkbook (workbook: XSSFWorkbook) (filePath: string) =
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


// #r "nuget: Npgsql.FSharp, 5.7.0"
// #r "nuget: SqlHydra.Query, 2.2.1";;
open Npgsql.FSharp

let connectionString = "postgres://postgres:XntSrCoEEZtiacZrx2m7jR5htEoEfYyoKncfhNmnPrLqPzxXTU5nxM@192.168.64.4:5432/lynx"
let qtestodelete = connectionString |> Sql.connect |> Sql.query "select * from lynx_sipnote where id = 27555;" |> Sql.execute (fun (read: RowReader) -> read.text "note")

type OIBRow = {
    LastName: string;
    FirstName: string;
    MiddleName: string option;
    ATDevices: bool;
    Orientation: bool;
    DLS: bool;
    Communications: bool;
    Advocacy: bool;
    Counseling: bool;
    Information: bool;
    Support: bool;
    PlanName: string;
    ATOutcomes: string option;
    CommunityPlanProgress: string option;
    ILSOutcomes: string option;
    LivingPlanProgress: string option;
    NoteDate: System.DateOnly
}

let queryColumns =
 "c.last_name, c.first_name, c.middle_name, n.at_devices, n.orientation, n.dls, n.communications, n.advocacy, n.counseling, n.information, n.support, p.plan_name, p.at_outcomes, p.community_plan_progress, p.ila_outcomes, p.living_plan_progress, n.note_date"

let joins = """
     lynx_sipnote AS n
JOIN lynx_contact AS c ON n.contact_id = c.id
JOIN lynx_sipplan AS p ON n.sip_plan_id = p.id
"""

let baseSelect = "SELECT " + queryColumns + " FROM " + joins

// TODO 2023-11-28_2107
// Convert to function that takes a from and to date. The format matters
// so it should probably be a constrained type.
let whereClause = "WHERE n.note_date >= '2022-10-01'::date AND n.note_date < '2023-10-01'::date"

let groupByClause = "GROUP BY " + queryColumns

let orderByClause = "ORDER BY CONCAT(c.last_name, ', ', c.first_name)"

let query = $"{baseSelect} {whereClause} {groupByClause} {orderByClause}"

let q = connectionString |> Sql.connect |> Sql.query query

let exeReader (read: RowReader) : OIBRow =
    {
        LastName = read.text "last_name";
        FirstName = read.text "first_name";
        MiddleName = read.textOrNone "middle_name";
        ATDevices = read.bool "at_devices";
        Orientation = read.bool "orientation";
        DLS = read.bool "dls";
        Communications = read.bool "communications";
        Advocacy = read.bool "advocacy";
        Counseling = read.bool "counseling";
        Information = read.bool "information";
        Support = read.bool "support";
        PlanName = read.text "plan_name";
        ATOutcomes = read.textOrNone "at_outcomes";
        CommunityPlanProgress = read.textOrNone "community_plan_progress";
        ILSOutcomes = read.textOrNone "ila_outcomes";
        LivingPlanProgress = read.textOrNone "living_plan_progress";
        NoteDate = read.dateOnly "note_date"
    }

let res = q |> Sql.execute exeReader

// === SqlHydra EXPERIMENTS ===
// User ID=postgres;Password=XntSrCoEEZtiacZrx2m7jR5htEoEfYyoKncfhNmnPrLqPzxXTU5nxM;Host=192.168.64.4;Port=5432;Database=lynx;

// $ dotnet fsi
// Microsoft (R) F# Interactive version 12.8.0.0 for F# 8.0
// Copyright (c) Microsoft Corporation. All Rights Reserved.

// For help type #help;;

// >  #r "nuget: Npgsql.FSharp, 5.7.0";;
// [Loading /Users/toraritte/.packagemanagement/nuget/Cache/697d8ca5b71fe39e0b2bf72bb58c700d58b82d6d086bcfc1fa356cce2708e407.fsx]
// module FSI_0003.
//        697d8ca5b71fe39e0b2bf72bb58c700d58b82d6d086bcfc1fa356cce2708e407

// > #r "nuget: SqlHydra.Query, 2.2.1";;
// [Loading /Users/toraritte/.packagemanagement/nuget/Cache/177be160dcb44a4a927d2619eda16eee3526dd5810c7ef37a6d4f9fd4544ce0d.fsx]
// module FSI_0002.
//        177be160dcb44a4a927d2619eda16eee3526dd5810c7ef37a6d4f9fd4544ce0d

// > #r "nuget: SqlHydra.Cli, 2.3.0";;

// /Users/toraritte/dev/clones/dotNET/slate-excel-reports/stdin(1,1): error FS0999: /Users/toraritte/.packagemanagement/nuget/Projects/85296--b0cee205-014b-423f-951e-e8bd674cb3f1/Proje
// ct.fsproj : error NU1202: Package SqlHydra.Cli 2.3.0 is not compatible with net8.0 (.NETCoreApp,Version=v8.0). Package SqlHydra.Cli 2.3.0 supports:

// > #r "nuget: SqlHydra.Cli, 2.3.1";;

// /Users/toraritte/dev/clones/dotNET/slate-excel-reports/stdin(1,1): error FS0999: /Users/toraritte/.packagemanagement/nuget/Projects/85296--b0cee205-014b-423f-951e-e8bd674cb3f1/Proje
// ct.fsproj : error NU1202: Package SqlHydra.Cli 2.3.1 is not compatible with net8.0 (.NETCoreApp,Version=v8.0). Package SqlHydra.Cli 2.3.1 supports:
// ====================

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