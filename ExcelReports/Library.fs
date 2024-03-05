module ExcelReports.OIB

(*
#r "nuget: NPOI, 2.6.2";;
#load "ExcelReports/OIBTypes.fs";;
open ExcelReports.OIBTypes;;

#load "ExcelReports/ExcelFunctions.fs";;
open ExcelReports.ExcelFunctions;;

#r "nuget: Npgsql.FSharp, 5.7.0";;
#load "ExcelReports/LynxData.fs";;
open ExcelReports.LynxData;;

#load "ExcelReports/Library.fs";;
open ExcelReports.OIB;;

-- oneliner:
fsi.ShowDeclarationValues <- false;;
#r "nuget: NPOI, 2.6.2";; #load "ExcelReports/OIBTypes.fs";; open ExcelReports.OIBTypes;; #load "ExcelReports/ExcelFunctions.fs";; open ExcelReports.ExcelFunctions;; #r "nuget: Npgsql.FSharp, 5.7.0";; #load "ExcelReports/LynxData.fs";; open ExcelReports.LynxData;; #load "ExcelReports/Library.fs";; open ExcelReports.OIB;;

let conn2 = "postgres://postgres:password@192.168.64.4:5432/lynx";;

generateQuarterlyReport conn2 OIB_7OB    2023 Q2 "dev";;
generateQuarterlyReport conn2 OIB_Non7OB 2023 Q2 "dev";;

generateQuarterlyReport conn  OIB_Non7OB 2023 Q2 "prod";;
generateQuarterlyReport conn  OIB_7OB    2023 Q2 "prod";;

let q = quarterlyReportQuery conn OIB_7OB Q2 2023;;
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

// TODO ok for now, but may need to be replaced with
//      a more meaningful type
type Client = string
type ClientOIBRows = (Client * OIBReportRow seq)
type OIBRowsGroupedAndOrderedByClientName = ClientOIBRows seq

let getClientName (row: QuarterlyReportQueryRow): OIBReportParseResult =
    let middleName = Option.defaultValue "" row.contact_middle_name
    let firstAndLastNames =
        ( row.contact_last_name
        , row.contact_first_name
        )
    match firstAndLastNames with
    | (None, _)
    | (_, None) ->
        Error $"Client name is missing in LYNX. (Contact ID: {row.contact_id})"
    | (Some last, Some first) ->
        let name = (last.Trim() + ", " + first.Trim() + " " + middleName.Trim())
        Ok (ClientName name)

let getIndividualsServed
    (row: QuarterlyReportQueryRow)
    (grantYearStart: System.DateOnly)
    : OIBReportParseResult =

    match row.intake_intake_date with
    | None ->
        Error "Intake date is not set in LYNX."
    | Some intakeDate ->
        match (intakeDate < grantYearStart) with
        | true  -> Ok PriorCase
        | false -> Ok NewCase

// Why `grantYearEnd`?
// Because a person may become 55 during the grant year.
let getAgeAtApplication
    (row: QuarterlyReportQueryRow)
    (reportType: QuarterlyOIBReportType)
    (grantYearEnd: System.DateOnly)
    : OIBReportParseResult =

    match row.intake_birth_date with
    | None ->
        Error $"Birth date is not set in LYNX. (Contact ID: {row.contact_id})"
    | Some (birthDate: System.DateOnly) ->
        match (birthDate > grantYearEnd) with
        | true ->
            Error $"Invalid date of birth: {birthDate}. (Contact ID: {row.contact_id})"
        | false ->
            let grantYearEndInDays = grantYearEnd.DayNumber
            let birthDateInDays = birthDate.DayNumber
            // Sloppy accomodation for leap years
            let age = (float (grantYearEndInDays - birthDateInDays)) / 365.25

            let ageAtApplication: AgeAtApplication =
                match age with
                // NOTE 2023-12-10_2006
                // There  are  (or will  be) age brackets  younger than
                // 55, but that is probably a clerical error as the SIP
                // program is only for individuals above 55.
                | _ when age < 25 -> AgeBracket18To24
                | _ when age < 35 -> AgeBracket25To34
                | _ when age < 45 -> AgeBracket35To44
                | _ when age < 55 -> AgeBracket45To54
                | _ when age < 65 -> AgeBracket55To64
                | _ when age < 75 -> AgeBracket65To74
                | _ when age < 85 -> AgeBracket75To84
                |              _  -> AgeBracket85AndOlder
                    // Error $"Age below 55. (DOB: {birthDate}, Contact ID: {row.contact_id})"

            let is7OB = reportType = OIB_7OB

            match (age, is7OB) with
            | (a, true) when a < 55 ->
                Error $"Age below 55. (DOB: {birthDate}, Contact ID: {row.contact_id})"
            | (a, false) when a >= 55 ->
                Error $"Age 55 or older. (DOB: {birthDate}, Contact ID: {row.contact_id})"
            | _ ->
                Ok ageAtApplication

// Takes type representation (i.e., `System.Type`) of a discriminated union with only case names, and tries to match it with a string supplied in the `match` clause.

(*
    Partial  active  pattern  for  converting  a  string
    from  a LYNX  database column  to a  case of  an OIB
    discriminated union type  in `OIBTypes.fs` (the type
    argument  also  has  to implement  the  `IOIBType`
    interface).

    Returns: on match -> IOIBType
             no match -> string

    For example:

        open OIBTypes
        let genderType = typeof<Gender>

        match "Female" with
        | OIBValue genderType matchedCaseIfAny ->
            matchedCaseIfAny // : IOIBType
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
    (oibType: System.Type) // ------*
    (valueToMatch: string)       // |
    : IOIBType option            // |
    =                            // |
                                 // V
    // Only able to check it at runtime (AFAIK).
    if (not <| typeof<IOIBType>.IsAssignableFrom(oibType))
    then failwith $"Type {oibType.FullName} does not implement the `IOIBType` interface."

    let caseToTuples (caseInfo: UnionCaseInfo) =
        let unionCase =
            FSharpValue.MakeUnion(caseInfo, [||]) :?> IOIBType
        ( (string unionCase)
        , unionCase
        )

    let valueMap=
        oibType
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
    (oibType: System.Type)
    (returnIfMatchFails: OIBReportParseResult option)
    (valueToMatch: string option)
    : OIBReportParseResult
    =

    match valueToMatch with
    | Some v ->
        match v with
        | OIBValue oibType case -> Ok case
        | other ->
            ( (Error $"Value '{other}' in Lynx is not a valid OIB option.")
            , returnIfMatchFails
            )
            ||> Option.defaultValue
    | None -> Error "Value is not set in LYNX."

// TODO Delete if not used anywhere
let getUnionType (case: obj) =
    let caseType = case.GetType()
    FSharpType.GetUnionCases(caseType).[0].DeclaringType

let getRace (row: QuarterlyReportQueryRow) : OIBReportParseResult =
    let raceType = typeof<Race>
    match row.intake_ethnicity with
    | Some "Two or More Races" -> Ok TwoOrMoreRaces
    | OIBCase raceType None result -> result

let getEthnicity (row: QuarterlyReportQueryRow) : OIBReportParseResult =
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
    match (row.intake_ethnicity, row.intake_other_ethnicity) with
    | (Some "Hispanic or Latino", _) -> Ok Yes
    | (None, Some _)                 -> Ok No
    | (_, Some "Hispanic or Latino") -> Ok Yes
    | (_, Some _)                    -> Ok No
    | (_, None)                      -> Ok No

let getDegreeOfVisualImpairment (row: QuarterlyReportQueryRow) : OIBReportParseResult =
    let degreeType = typeof<DegreeOfVisualImpairment>
    match row.intake_degree with
    // Historical LYNX options
    | Some "Light Perception Only" -> Ok LegallyBlind
    | Some "Low Vision" -> Ok SevereVisionImpairment
    | Some "Totally Blind (NP or NLP)" -> Ok TotallyBlind
    | OIBCase degreeType None result -> result

    // TODO 2023-12-11_1617
    // The "degree of visual impairment" in the OIB
    // report is mandatory.
    // TODO 2023-12-10_2009 Replace `failtwith`s with a visual cue in the OIB report
    // | None -> failwith $"Degree of visual impairment is null in LYNX. Client: {row.contact_id} {getClientName row}."
    // TODO 2023-12-10_2009 Replace `failtwith`s with a visual cue in the OIB report
    // | Some other -> failwith $"Degree of visual impairment '{other}' not in OIB report. Client: {row.contact_id} {getClientName row}."

// === HELPERS
let getColumn
    (columnType: System.Type)
    (nonOIBDefault: OIBReportParseResult option)
    (lynxColumn: string option)
    : OIBReportParseResult =

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
        * OIBReportParseResult option
        * string option
        )
    , OIBReportParseResult>()

let getColumnCached
    (columnType: System.Type)
    (nonOIBDefault: OIBReportParseResult option)
    (lynxColumn: string option)
    : OIBReportParseResult =
    let key = (columnType, nonOIBDefault, lynxColumn)
    match cache.TryGetValue(key) with
    | true, value -> value
    | _ ->
        let value = getColumn columnType nonOIBDefault lynxColumn
        cache.[key] <- value
        value

let boolOptsToResultYesNo (lynxColumns: bool option list) : OIBReportParseResult =

    let optTrueOrNone = function
        | Some b -> b
        | None -> false

    match (List.forall ((=) None) lynxColumns) with
    | true ->
        // There should probably be a descriptive
        // error type instead of a string.
        Error "Value not set in LYNX."
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
//   `row.intake_geriatric`       <->
//   `row.intake_stroke`          <->
//   `row.intake_seizure`         <->
//   `row.intake_migraine`        <->
//   `row.intake_heart`           <->
//   `row.intake_diabetes`        <->
//   `row.intake_dialysis`        <->
//   `row.intake_cancer`          <-> OtherImpairment
//   `row.intake_arthritis`       <->
//   `row.intake_high_bp`         <->
//   `row.intake_neuropathy`      <->
//   `row.intake_pain`            <->
//   `row.intake_asthma`          <->
//   `row.intake_musculoskeletal` <->
//   `row.intake_allergies        <->
//   `row.intake_dexterity`       <->
//
// In the case of the first 3, the presence of a value is crucial. The rest of the OIB columns are computed from multiple LYNX fields, so they can get away with a few missing values, but if all are missing, then then a human has to look into what is happening.
let getCognitiveImpairment (row: QuarterlyReportQueryRow) : OIBReportParseResult =
    [ row.intake_alzheimers
    ; row.intake_learning_disability
    ; row.intake_memory_loss
    ]
    |> boolOptsToResultYesNo

let getMentalHealthImpairment (row: QuarterlyReportQueryRow) : OIBReportParseResult =
    [ ( row.intake_mental_health
        |> Option.map (fun _ -> true)
      )
    ; row.intake_substance_abuse
    ]
    |> boolOptsToResultYesNo

let getOtherImpairment (row: QuarterlyReportQueryRow) : OIBReportParseResult =
    [ row.intake_geriatric
    ; row.intake_stroke
    ; row.intake_seizure
    ; row.intake_migraine
    ; row.intake_heart
    ; row.intake_diabetes
    ; row.intake_dialysis
    ; row.intake_cancer
    ; row.intake_arthritis
    ; row.intake_high_bp
    ; row.intake_neuropathy
    ; row.intake_pain
    ; row.intake_asthma
    ; row.intake_musculoskeletal
    ; ( row.intake_allergies
        |> Option.map (fun _ -> true)
      )
    ; row.intake_dexterity
    ]
    |> boolOptsToResultYesNo

let getTypeOfResidence (row: QuarterlyReportQueryRow) : OIBReportParseResult =
    let residenceType = typeof<TypeOfResidence>
    match row.intake_residence_type with
    // Historical LYNX options
    | Some "Community Residential" ->
        Ok SeniorIndependentLiving
    | Some "Skilled Nursing Care" ->
        Ok TypeOfResidence.NursingHome
    | Some "Assisted Living" ->
        Ok TypeOfResidence.AssistedLivingFacility
    | OIBCase residenceType None result ->
        result

let getReferrer (row: QuarterlyReportQueryRow) : OIBReportParseResult =
    let referrerType = typeof<SourceOfReferral>
    match row.intake_referred_by with
    // Historical LYNX options
    | Some "DOR" -> Ok StateVRAgency
    | OIBCase referrerType None result ->
        result


// ONLY DELETE AFTER THE HISTORICAL NOTE ARE MOVED TO THE DOCS!
// let getRace (row: QuarterlyReportQueryRow) : OIBReportParseResult =
//     let raceType = typeof<Race>
//     match row.intake_ethnicity with
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

let mapToDemographicsRow
    (reportType: QuarterlyOIBReportType)
    (grantYearStart: System.DateOnly)
    (grantYearEnd:   System.DateOnly)
    (row: QuarterlyReportQueryRow)
    : OIBReportRow =

    [
      ( "A", (getClientName row) )
    ; ( "B", getIndividualsServed row            grantYearStart )
    ; ( "E", getAgeAtApplication  row reportType grantYearEnd )
    ; ( "J", getColumnCached typeof<Gender> None row.intake_gender )
    // ; ( "N", getColumnCached typeof<Race> None row.intake_ethnicity )
    ; ( "N", getRace row )
    ; ( "V", getEthnicity row )
    ; ( "W", getDegreeOfVisualImpairment row )
    ; ( "AA", getColumnCached typeof<MajorCauseOfVisualImpairment> (Some <| Ok OtherCausesOfVisualImpairment) row.intake_eye_condition )
    ; ( "AG", boolOptsToResultYesNo [ row.intake_hearing_loss ] )
    ; ( "AH", boolOptsToResultYesNo [ row.intake_mobility ] )
    ; ( "AI", boolOptsToResultYesNo [ row.intake_communication ] )
    ; ( "AJ", getCognitiveImpairment row )
    ; ( "AK", getMentalHealthImpairment row )
    ; ( "AL", getOtherImpairment row )
    ; ( "AM", getTypeOfResidence row )
    // ; ( "AS", getColumnCached typeof<SourceOfReferral> None row.intake_referred_by )
    ; ( "AS", getReferrer row )
    ; ( "BF", getColumnCached typeof<County> None row.mostRecentAddress_county )
    ]
    // // For troubleshooting (to be able to compare the rows with the transformed ones).
    // (demoColumns, row)

let groupLynxRowsByClientName (rows: QuarterlyReportQueryRow seq) =
    rows
    |> Seq.groupBy (
        fun row ->
            match (getClientName row) with
            | Ok clientName -> string clientName
            | Error str -> str
        )

let getTabData
    (toOIBRows: QuarterlyReportQueryRow -> OIBReportRow )
    (queryData: OIBQuarterlyReportData)
    : OIBRowsGroupedAndOrderedByClientName =

    queryData.lynxData
    |> Seq.map (fun (row: ISQLQueryColumnable) ->
        row :?> QuarterlyReportQueryRow
     )
    |> groupLynxRowsByClientName
    |> Seq.sortBy fst
    |> Seq.map (fun (client: Client, lynxRows) ->
        lynxRows
        |> Seq.map toOIBRows
        |> fun oibRows ->
            (client, oibRows)
       )

let toReportColumn
    ( (colName: string
      , result: OIBReportParseResult
      ) : OIBReportColumn
    )
    : ReportColumn =

    let convertedResult: ReportCell =
        result
        |> Result.map (fun x -> x :?> System.IFormattable)
    (colName, convertedResult)

let toReportRow (oibRow: OIBReportRow) : ReportRow =
    oibRow
    |> List.map toReportColumn

let getDemographics
    (reportType: QuarterlyOIBReportType)
    (queryData: OIBQuarterlyReportData)
    : ReportSheetData
    =

    queryData
    |> getTabData
        ( mapToDemographicsRow
            reportType
            queryData.grantYearStart
            queryData.grantYearEnd
        )
    |> Seq.map (fun ((_clientName, oibRows): ClientOIBRows) ->
        oibRows
        |> Seq.map toReportRow
        |> Seq.distinct
        // All  `ReportRow`s   should  be the same  for a
        // given  client, so  if  this crashes,  it means  that
        // there is an issue with the LYNX data.
        //
        // TODO Wrap  it  in  `try..catch`  and  return  a
        //      more meaningful error  conveying the above
        //      message. Or just replace it with a `match`
        //      for a one element list.
        |> Seq.exactlyOne
    )

// ---SERVICES---------------------------------------------------------

// type ReportColumn = string * OIBReportParseResult
// type ReportRow = ReportColumn list

let getOutcomes (row: QuarterlyReportQueryRow) : OIBReportParseResult =
    match row.plan_living_plan_progress with
    | Some "Plan complete, no difference in ability to maintain living situation" ->
        Ok Maintained
    | Some "Plan complete, feeling more confident in ability to maintain living situation" ->
        Ok Increased
    | Some "Plan complete, feeling less confident in ability to maintain living situation" ->
        Ok Decreased
    | other ->
        Error $"Outcome needs to be set in LYNX or 'Case Status' (column V) needs to be 'Pending'."
    // Why no `NotAssessed`? See `case_status_conundrum` TODO below.

let getEmploymentOutcome (row: QuarterlyReportQueryRow) : OIBReportParseResult =
    let raceType = typeof<EmploymentOutcomes>
    match row.plan_employment_outcomes with
    | Some "Not Interested in Employment"    -> Ok NotInterested
    | Some "Less Likely to Seek Employment"  -> Ok LessLikely
    | Some "Unsure about Seeking Employment" -> Ok Unsure
    | Some "More Likely to Seek Employment"  -> Ok MoreLikely
    // This should never happen, but can't hurt.
    | OIBCase raceType None result -> result

// let getPlanDate (row: QuarterlyReportQueryRow) : OIBReportParseResult =
//     match row.plan_plan_date with
//     | Some date ->
//         Ok (PlanDate date)
//     | None ->
//         Error "No plan date in LYNX."

let getPlanModified (row: QuarterlyReportQueryRow) : OIBReportParseResult =
    match row.plan_modified with
    | Some date ->
        Ok (PlanModified date)
    | None ->
        Error "LYNX: plan.modified is NULL."

let mapToServicesRow (row: QuarterlyReportQueryRow) : OIBReportRow =
    [
      ( "_", getPlanModified row)
    //   ( "_", getPlanDate row )
    // ; ( "_", (Ok (PlanId row.plan_id)))
        // ------------------------------
      // TODO Ask what is with these rows
    ; ( "B", (Ok No)) // VisionAssessment
    ; ( "C", (Ok No)) // SurgicalOrTherapeuticTreatment
      // --------------------------------
    ; ( "D", boolOptsToResultYesNo [ row.note_at_devices; row.note_at_services ] )
    ; ( "E", getColumnCached typeof<AssistiveTechnologyGoalOutcomes> None row.plan_at_outcomes )
    ; ( "J", boolOptsToResultYesNo [ row.note_orientation ] )
    ; ( "K", boolOptsToResultYesNo [ row.note_communications] )
    ; ( "L", boolOptsToResultYesNo [ row.note_dls] )
    ; ( "M", boolOptsToResultYesNo [ row.note_advocacy] )
    ; ( "N", boolOptsToResultYesNo [ row.note_counseling ] )
    ; ( "O", boolOptsToResultYesNo [ row.note_information ] )
    ; ( "P", boolOptsToResultYesNo [ row.note_services ] )
    ; ( "Q", getColumnCached typeof<IndependentLivingAndAdjustmentOutcomes> None row.plan_ila_outcomes )
    ; ( "U", boolOptsToResultYesNo [ row.note_support ] )
      // TODO "case_status_conundrum"
      //      `CaseStatus` affects `LivingSituationOutcomes` (column W) and `HomeAndCommunityInvolvementOutcomes` (column AA), so if it is always assumed to be `Assessed`, then there is no point in every mapping to `NotAssessed`.
    ; ( "V",  (Ok Assessed) ) // CaseStatus
    ; ( "W",  getOutcomes row ) // LivingSituationOutcomes
    ; ( "AA", getOutcomes row ) // HomeAndCommunityInvolvementOutcomes
      // TODO Add to LYNX first then here
    ; ( "AE", getEmploymentOutcome row )
    ]

// To distinguish it from the `ReportRow` (= `ReportColumn list`) type.
type SameOIBColumns = OIBReportColumn seq

let byPlanModified ((("_", planModifiedResult) :: _rest): OIBReportRow) =
    match planModifiedResult with
    | Ok (planModified: IOIBType) ->
        let (PlanModified dateOnly) = (planModified :?> PlanModified)
        dateOnly
    | Error _ ->
        System.DateTime(1,1,1)

let mergeServiceYesNoCells (yesNoA: YesOrNo) (yesNoB: YesOrNo) : YesOrNo =
    match (yesNoA, yesNoB) with
    | (Yes, _)
    | (_, Yes) -> Yes
    | (No, No) -> No

let mergeOIBColumns
    ((colNameA, resultA): OIBReportColumn)
    ((colNameB, resultB): OIBReportColumn)
    : OIBReportColumn =

    // This should never happen, but doesn't hurt to check.
    if (colNameA <> colNameB) then failwith "Column names do not match."

    match (resultA, resultB) with
    | Error str, Ok _
    | Ok _, Error str ->
        (colNameA, Error str)
    | Error strA, Error strB ->
        if (strA = strB)
        then (colNameA, Error strA)
        else (colNameA, Error (strA + "; " + strB))
    | Ok oibTypeA as okA, Ok oibTypeB ->
        // Both `IOIBType`s should be the same type,
        // so doesn't matter which one.
        match (box oibTypeA) with

          // Irrelevent which  type abbreviation
          // it is; the rules are the same. (See
          // `mergeServiceYesNoCells`.)
        | :? YesOrNo ->
            ((oibTypeA :?> YesOrNo)
            ,(oibTypeB :?> YesOrNo)
            )
            ||> (mergeServiceYesNoCells) :> IOIBType
            |> Ok
            |> (fun x -> (colNameA, x))

          // always the same; see TODO "case_status_conundrum"
        | :? CaseStatus

          // By  the  time  this  function  is  called,
          // all transposed  `ReportRow`s will  be ordered
          // by `PlanModified`  date, so the  first one
          // is the most recent.
        | :? IOIBOutcome ->
            (colNameA, okA)

let getServices (queryData: OIBQuarterlyReportData) : ReportSheetData =
    queryData
    |> getTabData mapToServicesRow
    |> Seq.map (fun (_clientName: Client, oibRows: OIBReportRow seq) ->
        oibRows
        |> Seq.sortBy byPlanModified
        |> Seq.rev
           // A client's many "service rows"  is "mushed" into one
           // row, this  way  each column can be specified a merge
           // strategy.
        |> Seq.transpose
           // The  first  element of  the "Services"  `ReportRow`  is
           // not a  real column;  it was only  needed to  get the
           // outcome columns (`IOIBOutcome`) ordered.
        |> fun (t: SameOIBColumns seq) -> t |> Seq.tail
        |> Seq.map (fun (t: SameOIBColumns) ->
               t
               |> Seq.reduce mergeOIBColumns
           )
        |> Seq.toList
        |> toReportRow
    )

(*
    | SCENARIO | RECEIVED | OUTCOME |
    |          | SERVICE  |   SET   |
    |----------+----------+---------|
    |     1    |    Yes   |   Yes   | <- no brainer
    |     2    |    Yes   |   No    | <- disallowed
    |     3    |    No    |   Yes   | <- allowed with condition (see below)
    |     4    |    No    |   No    | <- no brainer
    |----------+----------+---------|

    Scenario  2  is  allowed,  if the  number of clients
    having received  services is higher than  the number
    of clients  having been assessed. No  checks needed,
    because once  the highlighted scenario 3  errors are
    corrected,  then  the  this  condition  will  always
    stand.
*)
let checkATServicesAndOutcome (row: OIBReportRow) : OIBReportRow =

    // Not  total  on  purpose:  it  should only  be called
    // with `Ok` `ParseResult`s; if  called with `Error` then
    // that is a bug.
    let okToError
        ( ( (letter: ColumnName)
          , (Ok (oibType: IOIBType): OIBReportParseResult)
          ): OIBReportColumn
        )
        : OIBReportColumn
        =
        (letter, Error (string oibType))

    let findColumn (letter: ColumnName) (row: OIBReportRow) : OIBReportColumn =
        row
        |> List.find (
               function
               | ((letter': ColumnName, _): OIBReportColumn) -> letter = letter'
           )

    let replaceColumns
        (row: OIBReportRow)
        (replacements: OIBReportColumn list)
        : OIBReportRow =

        let replacementLetters: ColumnName list =
            replacements |> List.map (fun (letter, _) -> letter)

        let needsReplacement (letter: ColumnName) =
            List.contains letter replacementLetters

        row
        |> List.map (
               function
               | ((letter: ColumnName, _): OIBReportColumn) when needsReplacement(letter) ->
                   findColumn letter replacements
               | otherColumn -> otherColumn
           )

    let atServiceYesNo = (findColumn "D" row)
    let      atOutcome = (findColumn "E" row)

    let affectedColumns =
        [ atServiceYesNo
        ;      atOutcome
        ]

    let resultsToCompare =
        affectedColumns
        |> List.map snd

    let replacementColumns : OIBReportColumn list =
        match resultsToCompare with
        | [ Ok no; outcomeResult ]
            when no = (No :> IOIBType) ->

            match outcomeResult with
            | Ok outcome
                when outcome = (NotAssessed :> IOIBType) ->
                []
            | _ ->
                affectedColumns
                |> List.map okToError
        | _ -> []

    replaceColumns row replacementColumns

let generateQuarterlyReport
    (connectionString: string)
    (reportType: QuarterlyOIBReportType)
    (year: int)
    (quarter: Quarter)
    (outPathSuffix: string)
    : unit
    =

    let oibData: OIBQuarterlyReportData =
        quarterlyReportQuery connectionString reportType quarter year

    // hard-coding it for now as it is not expected to change
    let templatePath =
        if (reportType = OIB_7OB)
        then "templates/20231011_protected_7-OB_Report-Data-Collection-Tool_V2.xlsx"
        else "templates/20231011_protected_Non-7-OB_Report-Data-Collection-Tool-for-under-OIB-age_V2.xlsx"

    let oaDate = System.DateTime.Now.ToOADate().ToString()
    let outPath =
        sprintf "%d_%s_%s_%s_%s.xlsx"
            year
            (string quarter)
            (string reportType)
            oaDate
            outPathSuffix

    let cellTransforms =
        // Sometimes using a previously generated report as a template, and error highlights need to be cleared - except for "Case Status" (column V) as it has a "default" color set by DOR.
        // TODO: implement setting "Case Status" and not setting it to "Assessed" indiscriminately.
        [ fun (cell: NPOI.SS.UserModel.ICell) ->
            if (cell.Address.Column <> 21)
            then resetCellColor cell
        ]

    openExcelFileWithNPOI templatePath
    |> populateSheet (getDemographics reportType oibData) 3 7 cellTransforms
    |> populateSheet (getServices oibData) 4 7 cellTransforms
    |> saveWorkbook outPath

let generateAssignmentReport
    (connectionString: string)
    (outPath: string)
    : unit
    =

    let templatePath = "templates/lynx_assignment_report_v3.xlsx"

    let query = assignmentReportQuery connectionString
    let query7OB = query OIB_7OB
    let queryNon7OB = query OIB_Non7OB

    let optionToReportCell (opt: 'a option) : ReportCell =
        match opt with
        | Some x -> Ok (GenericCellValue x)
        | None -> Error "Value not set in LYNX."

    let concatOpts (opt1: string option) (opt2: string option) : string option =
        Option.map2 (fun x y -> x + ", " + y) opt1 opt2

    let clientName (row: AssignmentQueryRow) : string option =
        concatOpts
            row.contact_last_name
            row.contact_first_name

    let instructorName (row: AssignmentQueryRow) : string option=
        concatOpts
            row.instructor_last_name
            row.instructor_first_name

    let assignedByName (row: AssignmentQueryRow) : string option=
        concatOpts
            row.assignedby_last_name
            row.assignedby_first_name

    let toReportRow
        (row: AssignmentQueryRow)
        (sipType: QuarterlyOIBReportType)
        : ReportRow
        =

        [
          ( "A", Ok (GenericCellValue sipType))
        ; ( "B", optionToReportCell row.assignment_assignment_date )
        ; ( "C", optionToReportCell (clientName row) )
        ; ( "D", optionToReportCell (instructorName row) )
        ; ( "E", optionToReportCell (assignedByName row) )
        ; ( "O", optionToReportCell row.assignment_assignment_status )
        ]

    let toSheetData
        (sipType: QuarterlyOIBReportType)
        (query: ISQLQueryColumnable list)
        : ReportSheetData
        =

        query
        |> Seq.map (fun row ->
            toReportRow (row :?> AssignmentQueryRow) sipType
        )

    let assignmentSheetData =
        toSheetData OIB_7OB query7OB
        |> Seq.append (toSheetData OIB_Non7OB queryNon7OB)

    // let addSIPType
    //     (sipType: QuarterlyOIBReportType)
    //     (query: ISQLQueryColumnable list)
    //     :
    //     =

    openExcelFileWithNPOI templatePath
    |> populateSheet assignmentSheetData 0 3 []
    |> saveWorkbook outPath

// TODO add functions to achieve the same as the fsi commands below, and make constraint checks pluggable as in the last comment block.

// let l = lynxQuery connectionString Q1 2023;;let ll = lynxQuery connectionString Q1 2022;; let d = getDemographics l;; let dd = getDemographics ll;; let s = getServices l;; let ss = getServices ll;;
// let o = openExcelFileWithNPOI "20231011_protected_7-OB_Report-Data-Collection-Tool_V2.xlsx";; populateSheet dd o 3;; populateSheet ss o 4;; saveWorkbook o "2022-2023.xlsx";;

(*
2 changes:
+ using a previously generated report as a template
+ "plugging in" the `checkATServicesAndOutcome` function to highlight additional errors

let o = openExcelFileWithNPOI "2023_Q1_rev6_012924-employment-filled.xlsx";; populateSheet d o 3;; populateSheet ( Seq.map checkATServicesAndOutcome s) o 4;; saveWorkbook o "2023_Q1_rev10_012924-employment-filled.xlsx";;
*)