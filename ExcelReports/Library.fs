module ExcelReports.OIB

open ExcelFunctions
open LynxData
open OIBTypes

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

let getClientName (lynxRow: LynxRow): ClientName =
    let clientName =
        lynxRow.contact_last_name + ", " +
        lynxRow.contact_first_name + " " +
        (Option.defaultValue "" lynxRow.contact_middle_name)
        |> fun str -> str.Trim() // if no middle name
    ClientName clientName

let getIndividualsServed (lynxRow: LynxRow) (grantYearStart: System.DateOnly) : IndividualsServed =
    let isIntakeBeforeGrantYear =
        lynxRow.intake_intake_date < grantYearStart
    match isIntakeBeforeGrantYear with
    | true ->  PriorCase
    | false -> NewCase

let getAgeAtApplication (lynxRow: LynxRow) (grantYearStart: System.DateOnly) : AgeAtApplication =
    let grantYearStartInDays = grantYearStart.DayNumber
    let birthDateInDays = lynxRow.intake_birth_date.DayNumber
    let ageAtApplicationInYears = (grantYearStartInDays - birthDateInDays) / 365
    match ageAtApplicationInYears with
    | _ when ageAtApplicationInYears < 55 -> failwith $"NOT IMPLEMENTED: Age below 55 not supported. Client: {lynxRow.contact_id} {getClientName lynxRow}."
    | _ when ageAtApplicationInYears < 65 -> AgeBracket55To64
    | _ when ageAtApplicationInYears < 75 -> AgeBracket65To74
    | _ when ageAtApplicationInYears < 85 -> AgeBracket75To84
    |                           _  -> AgeBracket85AndOlder

let getGender (lynxRow: LynxRow) : Gender =
    match lynxRow.intake_gender with
    | Some gender when gender = (toOIBString Male)   -> Male
    | Some gender when gender = (toOIBString Female) -> Female
    | Some other -> failwith $"NOT IMPLEMENTED: Gender '{other}' not supported. Client: {lynxRow.contact_id} {getClientName lynxRow}."
    | None   -> DidNotSelfIdentifyGender

let getRace (lynxRow: LynxRow) : Race =
    match lynxRow.intake_ethnicity with
    | Some race when race = (toOIBString NativeAmerican) ->
        NativeAmerican
    | Some race when race = (toOIBString Asian) ->
        Asian
    | Some race when race = (toOIBString AfricanAmerican) ->
        AfricanAmerican
    | Some race when race = (toOIBString PacificIslanderOrNativeHawaiian) ->
        PacificIslanderOrNativeHawaiian
    | Some race when race = (toOIBString White) ->
        White
    | Some race when race = (toOIBString TwoOrMoreRaces) ->
        TwoOrMoreRaces
    | Some race when race = (toOIBString DidNotSelfIdentifyRace) ->
        DidNotSelfIdentifyRace
    // HISTORICAL NOTE 2023-12-10_2222
    // LYNX  used  to  treat  the OIB  report's "Race"  and
    // "Ethnicity"  columns  in  one  field,  resulting  in
    // leaking  information.  The  workaround was  that  if
    // someone is  "Hispanic or Latino", then  the client's
    // race  was   set  to  "2  or   more  races",  instead
    // of  guessing.  (See  also  TODO  2023-12-02_2230  in
    // LynxData.fs; there may be more.)
    | Some "Hispanic or Latino" ->
        TwoOrMoreRaces
    // HISTORICAL NOTE 2023-12-10_2232
    // Another LYNX travesty: there was an "Other" option on
    // the  frontend that  had no corresponding value in the
    // OIB report.
    | Some "Other" ->
        DidNotSelfIdentifyRace
    | Some other ->
        failwith $"NOT IMPLEMENTED: Race '{other}' not supported. Client: {lynxRow.contact_id} {getClientName lynxRow}."
    | None ->
        DidNotSelfIdentifyRace

let getEthnicity (lynxRow: LynxRow) : HispanicOrLatino =
    // See HISTORICAL NOTEs 2023-12-10_2222 and
    // 2023-12-10_2232 in `getRace`.
    match (lynxRow.intake_ethnicity, lynxRow.intake_other_ethnicity) with
    | (Some "Hispanic or Latino", _) -> Yes
    | (None, Some _)                 -> No
    | (_, Some "Hispanic or Latino") -> Yes
    | (_, Some _)                    -> No
    | (_, None)                      -> No
    // | (Some this, Some that) -> failwith $"NOT IMPLEMENTED: Race {this}, Ethnicity {that}. Client: {lynxRow.contact_id} {getClientName lynxRow}."
    // | (None, Some that) -> failwith $"NOT IMPLEMENTED: Race None, Ethnicity {that}. Client: {lynxRow.contact_id} {getClientName lynxRow}."

let getDegreeOfVisualImpairment (lynxRow: LynxRow) : DegreeOfVisualImpairment =
    let degreeOfVisualImpairment =
        lynxRow.intake_degree_of_visual_impairment
        |> Option.defaultValue ""
    match degreeOfVisualImpairment with
    | "No Light Perception" -> NoLightPerception
    | "Light Perception Only" -> LightPerceptionOnly
    | "Low Vision" -> LowVision
    | "Blind" -> Blind
    | "Other" -> Other
    | "" -> NoLightPerception
    | other -> failwith $"NOT IMPLEMENTED: Degree of visual impairment '{other}' not supported. Client: {lynxRow.contact_id} {getClientName lynxRow}."

let fillDemographicsRow (lynxRow: LynxRow) (grantYearStart: System.DateOnly) : DemographicsRow =
    DemographicsRow
        ( (lynxRow |> getClientName),
          (getIndividualsServed lynxRow grantYearStart),
          (getAgeAtApplication  lynxRow grantYearStart),
          (lynxRow |> getGender),
          (lynxRow |> getRace),
          (lynxRow |> getEthnicity),
          (lynxRow |> getDegreeOfVisualImpairment),
          (lynxRow |> getMajorCauseOfVisualImpairment),
          (lynxRow |> getAgeRelatedImpairmentColumns),
          (lynxRow |> getTypeOfResidence),
          (lynxRow |> getSourceOfReferral),
          (lynxRow |> getCounty)
        )

//         * IndividualsServed            // "B7:D7"
//         * AgeAtApplication             // "E7:I7"
//         * Gender                       // "J7:M7"
//         // TODO 2023-12-10_1726
//         // Well, more like a note really, for when a
//         // LYNX query (see `lynxQuery`) "row" needs
//         // to be converted to a `DemographicsRow`.
//         // LYNX has the `lynx_intake` columns
//         // `ethnicity` and `other_ethnicity` that
//         // correspond to `Race` and `Ethnicity`
//         // respectively.
//         //
//         // The catch: `ethnicity` used to have all
//         // race options from the OIB report PLUS the
//         // ethnicity column (i.e., "Hispanic or
//         // Latino"), and `other_ethnicity` is mostly
//         // empty. So when `ethnicity` is "Hispanic
//         // or Latino", it means that `Race` will
//         // have to be set `TwoOrMoreRaces`... This
//         // has just been fixed in LYNX, but this has
//         // to be checked for backwards
//         // compatibility.
//         * Race                         // "N7:U7"
//         * Ethnicity                    // "V7"
//         * DegreeOfVisualImpairment     // "W7:Z7"
//         * MajorCauseOfVisualImpairment // "AA7:AF7"
//         * AgeRelatedImpairmentColumns  // "AG7:AL7"
//         * TypeOfResidence              // "AM7:AR7"
//         * SourceOfReferral             // "AS7:BE7"
//         * County                       // "BF7"
        // )
