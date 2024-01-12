module ExcelReports.OIB

(*
#r "nuget: NPOI, 2.6.2"
#load "ExcelReports/ExcelFunctions.fs";;
open ExcelReports.ExcelFunctions;;

#r "nuget: Npgsql.FSharp, 5.7.0";;
#load "ExcelReports/LynxData.fs";;
open ExcelReports.LynxData;;

#load "ExcelReports/OIBTypes.fs";;
open ExcelReports.OIBTypes;;

#load "ExcelReports/Library.fs";;
open ExcelReports.OIB;;
*)

open ExcelFunctions
open LynxData
open OIBTypes

open FSharp.Reflection

// The name of the opened Excel file is the version
// of this module.

// let xls = openExcelFileWithNPOI "20231011_protected_7-OB_Report-Data-Collection-Tool_V2.xlsx"

// Client names and demographic info has to be entered
// on  sheet "PART  III-DEMOGRAPHICS"  (4th sheet,  but
// NPOI indexing is zero-based) starting from cell "A7"
// (see "Instructions" sheet for details).
// let demoA7 = getCell xls 3 "A7"

// NOTE 2023-12-03_2348 subtract System.DateOnly instances
//     let subtractDateOnly (d1: System.DateOnly) (d2: System.DateOnly) = d1.Year - d2.Year
//
// a more elaborate way would be
//
//     let diff = d1.ToDateTime(System.TimeOnly.MinValue) - d2.ToDateTime(System.TimeOnly.MinValue)
//     let years = diff.Days / 365.25
//
// but it's not really worth it.

let getClientName (lynxRow: LynxRow): Result<IOIBString, string> =
    let middleName = Option.defaultValue "" lynxRow.contact_middle_name
    let firstAndLastNames =
        ( lynxRow.contact_last_name
        , lynxRow.contact_first_name
        )
    match firstAndLastNames with
    | (None, _)
    | (_, None) ->
        Error $"Client name is missing in LYNX. (Contact ID: {lynxRow.contact_id})"
    | (Some last, Some first) ->
        let name = (last.Trim() + ", " + first.Trim() + " " + middleName.Trim())
        Ok (ClientName name)

let getIndividualsServed
    (lynxRow: LynxRow)
    (grantYearStart: System.DateOnly)
    : Result<IOIBString, string> =

    match lynxRow.intake_intake_date with
    | None ->
        Error "Intake date is missing in LYNX."
    | Some intakeDate ->
        match (intakeDate < grantYearStart) with
        | true  -> Ok PriorCase
        | false -> Ok NewCase

// Why `grantYearEnd`?
// Because a person may become 55 during the grant year.
let getAgeAtApplication
    (lynxRow: LynxRow)
    (grantYearEnd: System.DateOnly)
    : Result<IOIBString, string> =

    match lynxRow.intake_birth_date with
    | None ->
        Error $"Birth date is missing in LYNX. (Contact ID: {lynxRow.contact_id})"
    | Some (birthDate: System.DateOnly) ->
        match (birthDate > grantYearEnd) with
        | true ->
            Error $"Invalid date of birth: {birthDate}. (Contact ID: {lynxRow.contact_id})"
        | false ->
            let grantYearEndInDays = grantYearEnd.DayNumber
            let birthDateInDays = birthDate.DayNumber
            // Sloppy accomodation for leap years
            let age = (float (grantYearEndInDays - birthDateInDays)) / 365.25

            match age with
            // NOTE 2023-12-10_2006
            // There  are  (or will  be) age brackets  younger than
            // 55, but that is probably a clerical error as the SIP
            // program is only for individuals above 55.
            | _ when age < 55 ->
                Error $"Age below 55. (DOB: {birthDate}, Contact ID: {lynxRow.contact_id})"
            | _ when age < 65 -> Ok AgeBracket55To64
            | _ when age < 75 -> Ok AgeBracket65To74
            | _ when age < 85 -> Ok AgeBracket75To84
            |              _  -> Ok AgeBracket85AndOlder

// Takes type representation (i.e., `System.Type`) of a discriminated union with only case names, and tries to match it with a string supplied in the `match` clause.

(*
    Partial  active  pattern  for  converting  a  string
    from  a LYNX  database column  to a  case of  an OIB
    discriminated union type  in `OIBTypes.fs` (the type
    argument  also  has  to implement  the  `IOIBString`
    interface).

    Returns: on match -> IOIBString
             no match -> string

    For example:

        open OIBTypes
        let genderType = typeof<Gender>

        match "Female" with
        | OIBValue genderType matchedCaseIfAny ->
            matchedCaseIfAny // : IOIBString
        | other ->
            // other : string
            failwith $"Value '{other}' in Lynx is not a valid OIB Gender option."

        match "lofa" with
        | OIBValue genderType matchedCaseIfAny -> toOIBString matchedCaseIfAny
        | other -> other


    > Why **partial** active pattern?
    > -------------------------------
    > Because
    >
    > 1. it needs to accept an OIB type argument
         to  be able  to  convert the  string to
         a specific union case, and
    >
    > 2. the  only  other  active pattern type that
         accepts  arguments  is  the  "one  choice"
         active pattern,  but that has to  return a
         concrete value; in the case of `OIBValue`,
         there is  a possibility  that there  is no
         match, so the  `option` type is necessary.
         (Could have just  thrown an exception, but
         we in  fact need to whether  it matches or
         not.)
*)
let (|OIBValue|_|)
    (iOIBStringType: System.Type) // active pattern argument
    (valueToMatch: string)
    =

    if (not <| typeof<IOIBString>.IsAssignableFrom(iOIBStringType))
    then failwith $"Type {iOIBStringType.FullName} does not implement the `IOIBString` interface."

    let caseToTuples (caseInfo: UnionCaseInfo) =
        let unionCase =
            FSharpValue.MakeUnion(caseInfo, [||]) :?> IOIBString
        ( (toOIBString unionCase)
        , unionCase
        )

    let valueMap=
        iOIBStringType
        |> FSharpType.GetUnionCases
        |> Array.map caseToTuples
        |> Map.ofArray

    Map.tryFind valueToMatch valueMap

(*
    `OIBCase` wraps `OIBValue` to
    + match `options string`  (instead of `string`)
    + provide a default value when `OIBValue` fails
      to  match  (if  `None`,  then  an  `Error` is
      returned)

    returns: Result

    For example:

        open OIBTypes
        let genderType = typeof<Gender>

        match (Some "Female") with
        | OIBCase genderType None result -> result;;

        match (Some "lofa") with
        | OIBCase genderType (Some (Error "lofa")) result -> result;;

    > Why **one-choice** active pattern?
    > ----------------------------------
    > Because  thanks to  the `returnIfMatchFails`,  it
    > always returns a value (in this case, a `Result`).
*)
let (|OIBCase|)
    (iOIBStringType: System.Type)
    (returnIfMatchFails: Result<IOIBString, string> option)
    (valueToMatch: string option) =

    match valueToMatch with
    | Some v ->
        match v with
        | OIBValue iOIBStringType case -> Ok case
        | other ->
            ( (Error $"Value '{other}' in Lynx is not a valid OIB option.")
            , returnIfMatchFails
            )
            ||> Option.defaultValue
    | None -> Error "Value is missing in LYNX."

// TODO Delete if not used anywhere
let getUnionType (case: obj) =
    let caseType = case.GetType()
    FSharpType.GetUnionCases(caseType).[0].DeclaringType

let getEthnicity (lynxRow: LynxRow) : Result<IOIBString, string> =
    // See HISTORICAL NOTEs 2023-12-10_2222 and
    // 2023-12-10_2232 in `getRace`.
    // -> FOLLOW-UP NOTE 2023-12-20_1156
    //    Decided  to go with the "codify the constraints that
    //    reflect how  things should be" approach,  instead of
    //    the "cater to buggy behaviors in the past" way.
    // -> FOLLOW-UP NOTE 2023-12-20_1210
    //    Haha,  lofty ideas  go brrr.  The `other_ethnicity`
    //    column is mostly `null`, and  I don't dare to change
    //    ca. 20000 rows and see  what breaks. Let's keep this
    //    noble task for the re-implementation of the LYNX.

    //                 OIB race                  OIB ethnicity
    //             ----------------          ----------------------
    match (lynxRow.intake_ethnicity, lynxRow.intake_other_ethnicity) with
    | (Some "Hispanic or Latino", _) -> Ok Yes
    | (None, Some _)                 -> Ok No
    | (_, Some "Hispanic or Latino") -> Ok Yes
    | (_, Some _)                    -> Ok No
    | (_, None)                      -> Ok No

let getDegreeOfVisualImpairment (lynxRow: LynxRow) : Result<IOIBString, string> =
    let degreeType = typeof<DegreeOfVisualImpairment>
    // NOTE "FS0025: Incomplete pattern match" warning
    //      The pattern  matches in  `OIBCase` active
    //      pattern are  exhaustive, but  the compiler
    //      has trouble figuring this out (plus, it is
    //      a nested active pattern, but `OIBValue` is
    //      also exhaustive).
    match lynxRow.intake_degree with
    // Historical LYNX options
    | Some "Light Perception Only" -> Ok LegallyBlind
    | Some "Low Vision" -> Ok SevereVisionImpairment
    | Some "Totally Blind (NP or NLP)" -> Ok TotallyBlind
    | OIBCase degreeType None result -> result

    // TODO 2023-12-11_1617
    // The "degree of visual impairment" in the OIB
    // report is mandatory.
    // TODO 2023-12-10_2009 Replace `failtwith`s with a visual cue in the OIB report
    // | None -> failwith $"Degree of visual impairment is null in LYNX. Client: {lynxRow.contact_id} {getClientName lynxRow}."
    // TODO 2023-12-10_2009 Replace `failtwith`s with a visual cue in the OIB report
    // | Some other -> failwith $"Degree of visual impairment '{other}' not in OIB report. Client: {lynxRow.contact_id} {getClientName lynxRow}."

// === HELPERS
let getColumn
    (columnType: System.Type)
    (nonOIBDefault: Result<IOIBString, string> option)
    (lynxColumn: string option)
    : Result<IOIBString, string> =

    // See NOTE "FS0025: Incomplete pattern match" warning above
    match lynxColumn with
    | OIBCase columnType nonOIBDefault result ->
        result

// Caching is needed because OIB types with many cases
// (such as  `County`) take 10s of  seconds to convert.
// (I  tried `Map`  at first,  but couldn't  get it  to
// work,  and `ConcurrentDictionary`  was suggested  by
// copilot.)
let cache =
    System.Collections.Concurrent.ConcurrentDictionary<
        ( System.Type
        * Result<IOIBString,string> option
        * string option
        )
    , Result<IOIBString,string>>()

let getColumnCached
    (columnType: System.Type)
    (nonOIBDefault: Result<IOIBString, string> option)
    (lynxColumn: string option)
    : Result<IOIBString, string> =
    let key = (columnType, nonOIBDefault, lynxColumn)
    match cache.TryGetValue(key) with
    | true, value -> value
    | _ ->
        let value = getColumn columnType nonOIBDefault lynxColumn
        cache.[key] <- value
        value

let hasImpairment (lynxColumns: bool option list) : Result<IOIBString, string> =

    let optTrueOrNone = function
        | Some b -> b
        | None -> false

    match lynxColumns with
    // Association of LYNX fields and "ager-related
    // impairment" OIB columns:
    //
    //   `intake_hearing_loss`  <-> HearingImpairment
    //   `intake_mobility`      <-> MobilityImpairment
    //   `intake_communication` <-> CommunicationImpairment
    //
    //   `intake_alzheimers`          <->
    //   `intake_learning_disability` <-> CognitiveImpairment
    //   `intake_memory_loss`         <->
    //   `intake_mental_health`   <->
    //   `intake_substance_abuse` <-> MentalHealthImpairment
    //   `lynxRow.intake_geriatric`       <->
    //   `lynxRow.intake_stroke`          <->
    //   `lynxRow.intake_seizure`         <->
    //   `lynxRow.intake_migraine`        <->
    //   `lynxRow.intake_heart`           <->
    //   `lynxRow.intake_diabetes`        <->
    //   `lynxRow.intake_dialysis`        <->
    //   `lynxRow.intake_cancer`          <-> OtherImpairment
    //   `lynxRow.intake_arthritis`       <->
    //   `lynxRow.intake_high_bp`         <->
    //   `lynxRow.intake_neuropathy`      <->
    //   `lynxRow.intake_pain`            <->
    //   `lynxRow.intake_asthma`          <->
    //   `lynxRow.intake_musculoskeletal` <->
    //   `lynxRow.intake_allergies        <->
    //   `lynxRow.intake_dexterity`       <->
    //
    // In the case of the first 3, the presence of a value is crucial. The rest of the OIB columns are computed from multiple LYNX fields, so they can get away with a few missing values, but if all are missing, then then a human has to look into what is happening.

    | _ when List.forall ((=) None) lynxColumns ->
        Error "Value is missing in LYNX."
    | _ ->
        lynxColumns
        |> List.tryFind optTrueOrNone
        |> (function
            | Some (Some true) -> Ok Yes
            | None -> Ok No
            // Yes, the  `match` is  not exhaustive  without these
            // cases,  but  they  will  never be  needed  based  on
            // `tryFind`'s output. (Or shouldn't be, for that matter;
            // and if they are, then the this will crash right away,
            // so that's good.)
            //
            // | Some (Some false) -> Ok No
            // | Some (None) -> Error "Value is missing in LYNX."
        )

let getCognitiveImpairment (lynxRow: LynxRow) : Result<IOIBString, string> =
    [ lynxRow.intake_alzheimers
    ; lynxRow.intake_learning_disability
    ; lynxRow.intake_memory_loss
    ]
    |> hasImpairment

let getMentalHealthImpairment (lynxRow: LynxRow) : Result<IOIBString, string> =
    [ ( lynxRow.intake_mental_health
        |> Option.map (fun _ -> true)
      )
    ; lynxRow.intake_substance_abuse
    ]
    |> hasImpairment

let getOtherImpairment (lynxRow: LynxRow) : Result<IOIBString, string> =
    [ lynxRow.intake_geriatric
    ; lynxRow.intake_stroke
    ; lynxRow.intake_seizure
    ; lynxRow.intake_migraine
    ; lynxRow.intake_heart
    ; lynxRow.intake_diabetes
    ; lynxRow.intake_dialysis
    ; lynxRow.intake_cancer
    ; lynxRow.intake_arthritis
    ; lynxRow.intake_high_bp
    ; lynxRow.intake_neuropathy
    ; lynxRow.intake_pain
    ; lynxRow.intake_asthma
    ; lynxRow.intake_musculoskeletal
    ; ( lynxRow.intake_allergies
        |> Option.map (fun _ -> true)
      )
    ; lynxRow.intake_dexterity
    ]
    |> hasImpairment

let getTypeOfResidence (lynxRow: LynxRow) : Result<IOIBString, string> =
    let residenceType = typeof<TypeOfResidence>
    match lynxRow.intake_residence_type with
    // Historical LYNX options
    | Some "Community Residential" ->
        Ok SeniorIndependentLiving
    | Some "Skilled Nursing Care" ->
        Ok TypeOfResidence.NursingHome
    | Some "Assisted Living" ->
        Ok TypeOfResidence.AssistedLivingFacility
    | OIBCase residenceType None result ->
        result

// ONLY DELETE AFTER THE HISTORICAL NOTE ARE MOVED TO THE DOCS!
// let getRace (lynxRow: LynxRow) : Result<IOIBString, string> =
//     let raceType = typeof<Race>
//     match lynxRow.intake_ethnicity with
//       // HISTORICAL NOTE 2023-12-10_2222
//       // LYNX  used  to  treat  the OIB  report's "Race"  and
//       // "Ethnicity"  columns  in  one  field,  resulting  in
//       // leaking  information.  The  workaround was  that  if
//       // someone is  "Hispanic or Latino", then  the client's
//       // race  was   set  to  "2  or   more  races",  instead
//       // of  guessing.  (See  also  TODO  2023-12-02_2230  in
//       // LynxData.fs; there may be more.)
//       // HISTORICAL NOTE 2023-12-10_2232
//       // There was an "Other" option on the frontend that had
//       // no corresponding value in the OIB report.
//       // -> FOLLOW-UP NOTE 2023-12-20_1145
//       //    Temporarily overriding the HISTORICAL NOTE  above
//       //    to test `Error` results.
//       //
//       // | Some "Hispanic or Latino" -> Ok TwoOrMoreRaces
//     | Some v ->
//         match v with
//         | OIBValue raceType race -> Ok race
//         | other -> Error $"Value '{other}' in Lynx is not a valid OIB option."
//     | None ->
//         Error "Value is missing in LYNX."

type DemographicsRow = (string * Result<IOIBString, string>) list

let createDemographicsRow
    (grantYearStart: System.DateOnly)
    (grantYearEnd:   System.DateOnly)
    (lynxRow: LynxRow)
    : DemographicsRow =
    // let demoColumns =
    [ ( "A", getClientName lynxRow )
    ; ( "B", getIndividualsServed lynxRow grantYearStart )
    ; ( "E", getAgeAtApplication  lynxRow grantYearEnd )
    ; ( "J", getColumnCached typeof<Gender> None lynxRow.intake_gender )
    ; ( "N", getColumnCached typeof<Race> None lynxRow.intake_ethnicity )
    ; ( "V", getEthnicity lynxRow )
    ; ( "W", getDegreeOfVisualImpairment lynxRow )
    ; ( "AA", getColumnCached typeof<MajorCauseOfVisualImpairment> (Some <| Ok OtherCausesOfVisualImpairment) lynxRow.intake_eye_condition )
    ; ( "AG", hasImpairment [ lynxRow.intake_hearing_loss ] )
    ; ( "AH", hasImpairment [ lynxRow.intake_mobility ] )
    ; ( "AI", hasImpairment [ lynxRow.intake_communication ] )
    ; ( "AJ", getCognitiveImpairment lynxRow )
    ; ( "AK", getMentalHealthImpairment lynxRow )
    ; ( "AL", getOtherImpairment lynxRow )
    ; ( "AM", getTypeOfResidence lynxRow )
    ; ( "AS", getColumnCached typeof<SourceOfReferral> None lynxRow.intake_referred_by )
    ; ( "BF", getColumnCached typeof<County> None lynxRow.mostRecentAddress_county )
    ]
    // // For troubleshooting (to be able to compare the rows with the transformed ones).
    // (demoColumns, lynxRow)

let getDemographics (lynxData: LynxData) : DemographicsRow seq =
    let rowsGroupedByClient =
        lynxData.lynxQuery
        |> Seq.map (createDemographicsRow lynxData.grantYearStart lynxData.grantYearEnd)
        |> Seq.groupBy (function
            | ((_, Ok client) :: _) -> toOIBString client
            | ((_, Error e)   :: _) -> failwith e
            | ([]) -> failwith "empty row"
        )

    rowsGroupedByClient
    |> Seq.map (fun (_client, demoRows) ->
        let consolidatedRows =
            demoRows
            |> Seq.distinct
            |> Seq.toList
        match consolidatedRows with
        | [row] -> row
        | _ -> failwith "A client has non-unique demographic rows."
    )

let fillDemographicsRow (dRow: DemographicsRow) (rowNumber: string) (xlsx: XSSFWorkbook) =
    let errorColor = hexStringToRGB "#ffc096"
    dRow
    |> Seq.iter (
        fun (column, result) ->
            // let rowNum = string(i + 7)
            let cell = getCell xlsx 3 (column + rowNumber)
            let cellString =
                match result with
                | Ok oibValue ->
                    toOIBString oibValue
                | Error str ->
                    changeCellColor cell errorColor
                    "Error: " + str
            updateCell cell cellString
    )

let populateDemographicsTab (dRows: DemographicsRow seq) (xlsx: XSSFWorkbook) =
    dRows
    |> Seq.iteri (
        fun i row ->
         fillDemographicsRow row (string(i + 7)) xlsx
       )

// ---SERVICES---------------------------------------------------------
let mergeServiceRows rowA rowB =
    (rowA, rowB)
    ||> Seq.zip
    |> Seq.map (fun ((column, valueA), (_, valueB)) ->
        let shouldBeOneButIsIt =
            ["A"; "B"; "E"; "J"; "N"; "V"; "W"; "AA"; "AG"; "AH"; "AI"; "AJ"; "AK"; "AL"; "AM"; "AS"; "BF"]
        match valueA with
        | Ok _ -> (column, valueA)
        | Error _ -> (column, valueB)
    )
