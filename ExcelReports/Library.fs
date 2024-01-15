﻿module ExcelReports.OIB

(*
#r "nuget: NPOI, 2.6.2";;
#load "ExcelReports/ExcelFunctions.fs";;
open ExcelReports.ExcelFunctions;;

#r "nuget: Npgsql.FSharp, 5.7.0";;
#load "ExcelReports/LynxData.fs";;
open ExcelReports.LynxData;;

#load "ExcelReports/OIBTypes.fs";;
open ExcelReports.OIBTypes;;

#load "ExcelReports/Library.fs";;
open ExcelReports.OIB;;

-- oneliner:
#r "nuget: NPOI, 2.6.2";; #load "ExcelReports/ExcelFunctions.fs";; open ExcelReports.ExcelFunctions;; #r "nuget: Npgsql.FSharp, 5.7.0";; #load "ExcelReports/LynxData.fs";; open ExcelReports.LynxData;; #load "ExcelReports/OIBTypes.fs";; open ExcelReports.OIBTypes;; #load "ExcelReports/Library.fs";; open ExcelReports.OIB;;
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

let boolOptsToResultYesNo (lynxColumns: bool option list) : Result<IOIBString, string> =

    let optTrueOrNone = function
        | Some b -> b
        | None -> false

    match (List.forall ((=) None) lynxColumns) with
    | true ->
        // There should probably be a descriptive
        // error type instead of a string.
        Error "Value is missing in LYNX."
    | false ->
        lynxColumns
        |> List.tryFind optTrueOrNone
        |> (function
            | Some (Some true) -> Ok Yes
            | None -> Ok No
            (*
                Yes, the  match is not exhaustive,  but the cases
                at the bottom will never happen:

                possible   |               |
                 inputs    | optTrueOrNone | tryFind optTrueOrNone
                -----------+---------------+----------------------
                None       | false         | None
                Some true  | true          | Some true
                Some false | false         | None
            *)
            // | Some (Some false) -> Ok No
            // | Some (None) -> Error "Value is missing in LYNX."
        )

// Association of LYNX fields and "age-related
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
let getCognitiveImpairment (lynxRow: LynxRow) : Result<IOIBString, string> =
    [ lynxRow.intake_alzheimers
    ; lynxRow.intake_learning_disability
    ; lynxRow.intake_memory_loss
    ]
    |> boolOptsToResultYesNo

let getMentalHealthImpairment (lynxRow: LynxRow) : Result<IOIBString, string> =
    [ ( lynxRow.intake_mental_health
        |> Option.map (fun _ -> true)
      )
    ; lynxRow.intake_substance_abuse
    ]
    |> boolOptsToResultYesNo

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
    |> boolOptsToResultYesNo

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

type OIBColumn = string * Result<IOIBString, string>
type OIBRow = OIBColumn list

let mapToDemographicsRow
    (grantYearStart: System.DateOnly)
    (grantYearEnd:   System.DateOnly)
    (lynxRow: LynxRow)
    : OIBRow =
    // let demoColumns =
    [ ( "A", getClientName lynxRow )
    ; ( "B", getIndividualsServed lynxRow grantYearStart )
    ; ( "E", getAgeAtApplication  lynxRow grantYearEnd )
    ; ( "J", getColumnCached typeof<Gender> None lynxRow.intake_gender )
    ; ( "N", getColumnCached typeof<Race> None lynxRow.intake_ethnicity )
    ; ( "V", getEthnicity lynxRow )
    ; ( "W", getDegreeOfVisualImpairment lynxRow )
    ; ( "AA", getColumnCached typeof<MajorCauseOfVisualImpairment> (Some <| Ok OtherCausesOfVisualImpairment) lynxRow.intake_eye_condition )
    ; ( "AG", boolOptsToResultYesNo [ lynxRow.intake_hearing_loss ] )
    ; ( "AH", boolOptsToResultYesNo [ lynxRow.intake_mobility ] )
    ; ( "AI", boolOptsToResultYesNo [ lynxRow.intake_communication ] )
    ; ( "AJ", getCognitiveImpairment lynxRow )
    ; ( "AK", getMentalHealthImpairment lynxRow )
    ; ( "AL", getOtherImpairment lynxRow )
    ; ( "AM", getTypeOfResidence lynxRow )
    ; ( "AS", getColumnCached typeof<SourceOfReferral> None lynxRow.intake_referred_by )
    ; ( "BF", getColumnCached typeof<County> None lynxRow.mostRecentAddress_county )
    ]
    // // For troubleshooting (to be able to compare the rows with the transformed ones).
    // (demoColumns, lynxRow)

let groupLynxRowsByClientName (rows: LynxQuery) =
    rows
    |> Seq.groupBy (
        fun lynxRow ->
            match (getClientName lynxRow) with
            | Ok clientName -> toOIBString clientName
            | Error str -> str
        )

let getTabData
    (lynxRowMapper: LynxRow -> OIBRow )
    (lynxData: LynxData)
    : (string * OIBRow seq) seq =

    lynxData.lynxQuery
    |> groupLynxRowsByClientName
    |> Seq.sortBy fst
    |> Seq.map (fun (clientName, lynxRows) ->
        lynxRows
        |> Seq.map lynxRowMapper
        |> fun oibRows -> (clientName, oibRows)
       )

let getDemographics (lynxData: LynxData) : OIBRow seq =
    lynxData
    |> getTabData (mapToDemographicsRow lynxData.grantYearStart lynxData.grantYearEnd)
    |> Seq.map (fun (_clientName, oibRows) ->
        oibRows
        |> Seq.distinct
        // All  `OIBRow`s   should  be the same  for a
        // given  client, so  if  this crashes,  it means  that
        // there is an issue with the LYNX data.
        //
        // TODO Wrap  it  in  `try..catch`  and  return  a
        //      more meaningful error  conveying the above
        //      message. Or just replace it with a `match`
        //      for a one element list.
        |> Seq.exactlyOne
    )

let fillRow (dRow: OIBRow) (rowNumber: string) (sheetNumber: int) (xlsx: XSSFWorkbook) =
    let errorColor = hexStringToRGB "#ffc096"
    dRow
    |> Seq.iter (
        fun ((column, result): OIBColumn) ->
            // let rowNum = string(i + 7)
            let cell = getCell xlsx sheetNumber (column + rowNumber)
            let cellString =
                match result with
                | Ok oibValue ->
                    toOIBString oibValue
                | Error str ->
                    changeCellColor cell errorColor
                    "Error: " + str
            updateCell cell cellString
    )

// let extractClientName (dRow: OIBRow) =
//     match dRow with
//     | (head :: _) ->
//         match (snd head) with
//         | Ok name -> toOIBString name
//         | Error str -> str
//     | _ -> failwith "Malformed demographics row."

let populateSheet (dRows: OIBRow seq) (xlsx: XSSFWorkbook) (sheetNumber: int) =
    dRows
    // |> Seq.sortBy extractClientName
    |> Seq.iteri (
        fun i row ->
         fillRow row (string(i + 7)) sheetNumber xlsx
       )

// ---SERVICES---------------------------------------------------------

// type OIBColumn = string * Result<IOIBString, string>
// type OIBRow = OIBColumn list

let getOutcomes (lynxRow: LynxRow) : Result<IOIBString, string> =
    let degreeType = typeof<DegreeOfVisualImpairment>
    match lynxRow.plan_living_plan_progress with
    | Some "Plan complete, no difference in ability to maintain living situation" -> Ok Maintained
    | Some "Plan complete, feeling more confident in ability to maintain living situation" -> Ok Increased
    | Some "Plan complete, feeling less confident in ability to maintain living situation" -> Ok Decreased
    | other -> Error $"Error: LYNX value: '{other}'."
    // Why no `NotAssessed`? See `case_status_conundrum` TODO below.

// let getPlanDate (lynxRow: LynxRow) : Result<IOIBString, string> =
//     match lynxRow.plan_plan_date with
//     | Some date ->
//         Ok (PlanDate date)
//     | None ->
//         Error "No plan date in LYNX."

let getPlanModified (lynxRow: LynxRow) : Result<IOIBString, string> =
    match lynxRow.plan_modified with
    | Some date ->
        Ok (PlanModified date)
    | None ->
        Error "LYNX: plan.modified is NULL."

let mapToServicesRow (lynxRow: LynxRow) : OIBRow =
    [
      ( "_", getPlanModified lynxRow)
    //   ( "_", getPlanDate lynxRow )
    // ; ( "_", (Ok (PlanId lynxRow.plan_id)))
        // ------------------------------
      // TODO Ask what is with these rows
    ; ( "B", (Ok No)) // VisionAssessment
    ; ( "C", (Ok No)) // SurgicalOrTherapeuticTreatment
      // --------------------------------
    ; ( "D", boolOptsToResultYesNo [ lynxRow.note_at_devices; lynxRow.note_at_services ] )
    ; ( "E", getColumnCached typeof<AssistiveTechnologyGoalOutcomes> None lynxRow.plan_at_outcomes )
    ; ( "J", boolOptsToResultYesNo [ lynxRow.note_orientation ] )
    ; ( "K", boolOptsToResultYesNo [ lynxRow.note_communications] )
    ; ( "L", boolOptsToResultYesNo [ lynxRow.note_dls] )
    ; ( "M", boolOptsToResultYesNo [ lynxRow.note_advocacy] )
    ; ( "N", boolOptsToResultYesNo [ lynxRow.note_counseling ] )
    ; ( "O", boolOptsToResultYesNo [ lynxRow.note_information ] )
    ; ( "P", boolOptsToResultYesNo [ lynxRow.note_services ] )
    ; ( "Q", getColumnCached typeof<IndependentLivingAndAdjustmentOutcomes> None lynxRow.plan_ila_outcomes )
    ; ( "U", boolOptsToResultYesNo [ lynxRow.note_support ] )
      // TODO "case_status_conundrum"
      //      `CaseStatus` affects `LivingSituationOutcomes` (column W) and `HomeAndCommunityInvolvementOutcomes` (column AA), so if it is always assumed to be `Assessed`, then there is no point in every mapping to `NotAssessed`.
    ; ( "V",  (Ok Assessed) ) // CaseStatus
    ; ( "W",  getOutcomes lynxRow )
    ; ( "AA", getOutcomes lynxRow )
      // TODO Add to LYNX first then here
    // ; ( "AE", getColumnCached typeof<EmploymentOutcomes> None lynxRow.plan_employment_outcomes )
    ]

// To distinguish it from the `OIBRow` (= `OIBColumn list`) type.
type TransposedOIBRow = OIBColumn seq

// let getResult (t: TransposedOIBRow) : Result<IOIBString, string> =
//     // Intentionally not  using  `Seq.tryFind`: if
//     // there are only `Error`s, then  something is
//     // very off and needs to be investigated.
//     let firstOk =
//         t
//         |> Seq.find (fun (_columnLetter, result) -> result = Ok Yes)
//     let (_columnLetter, result) = firstOk
//     result

// The current rule to consolidate `Result<YesOrNo, string>`s:
// 1. If there is a `Ok Yes`, then that is the result.
// 2. If  there is  an  `Error`,  return that  as the rest
//    are  `Ok No`s,  so  it may  be  possible that  after
//    fixing  the  error, the  result  will  be `Ok  Yes`.
//    (See  `boolOptsToResultYesNo`  for what  errors  are
//    possible.)
// 3. All elements are `Ok No` at this point.
//
// > WHY RETURN A SINGLE ERROR IN SCENARIO 2 AND NOT MERGING ERRORS WITH NOTE IDS?
// > -----------------------------------------------------------------------------
// > It  seems  to  be  too  much  effort for  too little
// > gain: LYNX columns that comprise the `YesOrNo` cells
// > on the "Services"  tab are of `bool  option` type in
// > `lynxQuery`  out  of  caution,  but so  far  I  have
// > not  seen any  nulls  (at the  moment,  that is  the
// > only  error  returned  by  `boolOptsToResultYesNo`).
// > Therefore, if there is an error (i.e., null) then it
// > can be  looked up  in the  clients records,  and the
// > others should show up as well.
let mergeServiceYesNoCells (t: TransposedOIBRow) : OIBColumn =
        t
        |> Seq.tryFind ( fun ((_colLetter, result): OIBColumn) -> result = Ok Yes )
        |> function
            | None ->
                let anyError =
                    t
                    |> Seq.tryFind (
                        function
                        | (_, Error _) -> true
                        | _ -> false)
                match anyError with
                | Some errorColumn -> errorColumn
                | None -> t |> Seq.head
            | Some okColumn  -> okColumn

// let byPlanDate ((("_", planDateResult) :: _rest): OIBRow) =
//     match planDateResult with
//     | Ok (planDate: IOIBString) ->
//         let (PlanDate dateOnly) = (planDate :?> PlanDate)
//         dateOnly
//     | Error _ ->
//         System.DateOnly(1,1,1)

let byPlanModified ((("_", planModifiedResult) :: _rest): OIBRow) =
    match planModifiedResult with
    | Ok (planModified: IOIBString) ->
        let (PlanModified dateOnly) = (planModified :?> PlanModified)
        dateOnly
    | Error _ ->
        System.DateTime(1,1,1)

// By  the  time  this  function  is  called,
// all transposed  `OIBRow`s will  be ordered
// by `PlanModified`  date, so the  first one
// is the most recent.
let mergeOutcomes (t: TransposedOIBRow) : OIBColumn =
    t
    // |> Seq.distinct
    |> Seq.head

let getServices (lynxData: LynxData) : OIBRow seq =
    lynxData
    |> getTabData mapToServicesRow
    |> Seq.map (fun (clientName, oibRows) ->
        oibRows
        |> Seq.sortBy byPlanModified
        |> Seq.rev
    //     |> fun x -> (clientName, x)
    // )
        |> Seq.transpose
           // The  first  element of  the "Services"  `OIBRow`  is
           // not a  real column;  it was only  needed to  get the
           // outcome columns (`IOIBOutcome`) ordered.
        |> fun t -> t |> Seq.tail
        |> Seq.map (fun (t: TransposedOIBRow) ->
            // let (Ok ioibString) = getResult t
            match (t |> Seq.head |> snd) with
            | Error str -> t |> Seq.head
            | Ok ioibString ->
                match (box ioibString) with
                // Irrelevent which  type abbreviation
                // it is; the rules are the same. (See
                // `mergeServiceYesNoCells`.)
                | :? YesOrNo -> mergeServiceYesNoCells t
                | :? IOIBOutcome -> mergeOutcomes t
                // It is fixed in `mapToServicesRow`.
                | :? CaseStatus as cs -> t |> Seq.head
           )
        |> Seq.toList
    )
