module ExcelReports.OIB

open ExcelFunctions
open LynxData

// The name of the opened Excel file is the version
// of this module.

let xls = openExcelFileWithNPOI "../20231011_protected_7-OB_Report-Data-Collection-Tool_V2.xlsx"

// Client names and demographic info has to be entered
// on  sheet "PART  III-DEMOGRAPHICS"  (4th sheet,  but
// NPOI indexing is zero-based) starting from cell "A7"
// (see "Instructions" sheet for details).
let demoA7 = getCell xls 3 "A7"

// To get list validation values quickly
let gv (cell: ICell) =
    let getWorkbook (cell: ICell) = cell.Sheet.Workbook :?> XSSFWorkbook
    let getSheetIndex (cell: ICell) = (getWorkbook cell).GetSheetIndex(cell.Sheet)

    ( findDataValidation cell
      |> Option.get).ValidationConstraint.Formula1
      |> fun str -> str.Replace("$",""
    )
    |> convertCellRangeToList
    |> List.map
        (fun address ->
            (address,
            (getCell 
                (getWorkbook cell) 
                (getSheetIndex cell)
                address).StringCellValue)
        )
    |> fun list -> ( ((findMergedRegion cell |> Option.get).FormatAsString()), list)

let gev (cell: ICell) =
    (findDataValidation cell |> Option.get).ValidationConstraint.ExplicitListValues

// NOTE 2023-12-03_2348 subtract System.DateOnly instances
//     let subtractDateOnly (d1: System.DateOnly) (d2: System.DateOnly) = d1.Year - d2.Year
//
// a more elaborate way would be
//
//     let diff = d1.ToDateTime(System.TimeOnly.MinValue) - d2.ToDateTime(System.TimeOnly.MinValue)
//     let years = diff.Days / 365.25
//
// but it's not really worth it.
