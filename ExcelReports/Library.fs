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

// NOTE 2023-12-03_2153
// F# supports this out of the box:
//
// + `nameof` (e.g., nameof MyUnion.Opt2)
// + `string` (e.g., string MyUnion.Opt2 or let a = Opt2; string a)
//
// open Microsoft.FSharp.Reflection
//
// let getUnionCases<'T> =
//     FSharpType.GetUnionCases(typeof<'T>)
//     |> Array.map (fun caseInfo -> caseInfo.Name)

type IndividualsServed = // individuals served
| NewCase
| PriorCase

type Demographics =
| IndividualsServed of IndividualsServed

type ListValidationValues =
| Demographics of Demographics

// TODO create the list validation types and then collect them in a union type, then replace 'a with it
//      Would this work?...
let listValidationTypeToString (validationValue: ListValidationValues) =
    match validationValue with
    | Demographics (IndividualsServed PriorCase) -> "Case open prior to Oct. 1"
    | Demographics (IndividualsServed NewCase)   -> "Case open between Oct. 1 - Sept. 30"
// ["Case open prior to Oct. 1"; "Case open between Oct. 1 - Sept. 30"]

// To get list validation values quickly
let gv (cell: ICell) =
    (findDataValidation cell
    |> Option.get).ValidationConstraint.Formula1
    |> fun str -> str.Replace("$","")
    |> convertCellRangeToList
    |> List.map (fun address -> (address, (getCell (cell.Sheet.Workbook :?> XSSFWorkbook) 3 address).StringCellValue) )
    |> fun list -> ( ((findMergedRegion cell |> Option.get).FormatAsString()), list)

// NOTE 2023-12-03_2348 subtract System.DateOnly instances
//     let subtractDateOnly (d1: System.DateOnly) (d2: System.DateOnly) = d1.Year - d2.Year
//
// a more elaborate way would be
//
//     let diff = d1.ToDateTime(System.TimeOnly.MinValue) - d2.ToDateTime(System.TimeOnly.MinValue)
//     let years = diff.Days / 365.25
//
// but it's not really worth it.