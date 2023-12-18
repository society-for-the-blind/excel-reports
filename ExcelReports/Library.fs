﻿module ExcelReports.OIB

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
    let getAge (lynxRow: LynxRow) (grantYearStart: System.DateOnly) =
        let grantYearStartInDays = grantYearStart.DayNumber
        let birthDateInDays = lynxRow.intake_birth_date.DayNumber
        // Sloppy accomodation for leap years
        (float (grantYearStartInDays - birthDateInDays)) / 365.25

    let age = getAge lynxRow grantYearStart
    match lynxRow.intake_birth_date with
    // NOTE 2023-12-10_2006
    // There are (or will be) age brackets younger
    // than 55, but that is probably a clerical
    // error as the SIP program is only for
    // individuals above 55.
    //
    // TODO 2023-12-10_2009 Replace `failtwith`s with a visual cue in the OIB report
    //
    | None -> failwith $"No birth date in LYNX. Client: {lynxRow.contact_id} {getClientName lynxRow}."
    | _ when age < 55 -> failwith $"Age below 55 not in OIB report. Client: {lynxRow.contact_id} {getClientName lynxRow}."
    | _ when age < 65 -> AgeBracket55To64
    | _ when age < 75 -> AgeBracket65To74
    | _ when age < 85 -> AgeBracket75To84
    |              _  -> AgeBracket85AndOlder

let getGender (lynxRow: LynxRow) : Gender =
    match lynxRow.intake_gender with
    | Some gender when gender = (toOIBString Male)   -> Male
    | Some gender when gender = (toOIBString Female) -> Female
    | Some other -> failwith $"Gender '{other}' not in OIB report. Client: {lynxRow.contact_id} {getClientName lynxRow}."
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
        failwith $"Race '{other}' not in OIB report. Client: {lynxRow.contact_id} {getClientName lynxRow}."
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
    // | (Some this, Some that) -> failwith $"Race {this}, Ethnicity {that}. Client: {lynxRow.contact_id} {getClientName lynxRow}."
    // | (None, Some that) -> failwith $"Race None, Ethnicity {that}. Client: {lynxRow.contact_id} {getClientName lynxRow}."

let getDegreeOfVisualImpairment (lynxRow: LynxRow) : DegreeOfVisualImpairment =
    match lynxRow.intake_degree with
    | Some degree when degree = (toOIBString TotallyBlind) ->
        TotallyBlind
    | Some "Totally Blind (NP or NLP)" ->
        TotallyBlind
    | Some degree when degree = (toOIBString LegallyBlind) ->
        LegallyBlind
    | Some degree when degree = (toOIBString SevereVisionImpairment) ->
        SevereVisionImpairment
    // Historical LYNX options
    | Some "Light Perception Only" ->
        LegallyBlind
    | Some "Low Vision" ->
        SevereVisionImpairment
    // TODO 2023-12-11_1617
    // The "degree of visual impairment" in the OIB
    // report is mandatory.
    // TODO 2023-12-10_2009 Replace `failtwith`s with a visual cue in the OIB report
    | None -> failwith $"Degree of visual impairment is null in LYNX. Client: {lynxRow.contact_id} {getClientName lynxRow}."
    // TODO 2023-12-10_2009 Replace `failtwith`s with a visual cue in the OIB report
    | Some other -> failwith $"Degree of visual impairment '{other}' not in OIB report. Client: {lynxRow.contact_id} {getClientName lynxRow}."

let getMajorCauseOfVisualImpairment (lynxRow: LynxRow) : MajorCauseOfVisualImpairment =
    match lynxRow.intake_eye_condition with
    | Some eyeCondition when eyeCondition = (toOIBString MacularDegeneration) ->
        MacularDegeneration
    | Some eyeCondition when eyeCondition = (toOIBString DiabeticRetinopathy) ->
        DiabeticRetinopathy
    | Some eyeCondition when eyeCondition = (toOIBString Glaucoma) ->
        Glaucoma
    | Some eyeCondition when eyeCondition = (toOIBString Cataracts) ->
        Cataracts
    | Some eyeCondition when eyeCondition = (toOIBString OtherCausesOfVisualImpairment) ->
        OtherCausesOfVisualImpairment
    // NOTE 2023-12-11_1920
    // Clients imported from the old system have all
    // kinds of entries because it didn't have a
    // dropdown, but a text field.
    | Some _ -> OtherCausesOfVisualImpairment
    | None   -> OtherCausesOfVisualImpairment

let getAgeRelatedImpairmentColumns (lynxRow: LynxRow) : AgeRelatedImpairmentColumns =

    // === HELPERS
    let hasImpairment (lynxColumns: bool list) : YesOrNo =
        match (List.contains true lynxColumns) with
        | true -> Yes
        | false -> No

    let optstringToBool (optstring: string option) : bool =
        match optstring with
        | Some _ -> true
        | None   -> false
    // `optstring |> Option.isNone |> not` is shorter, but more obscure
    // ===

    let getHearingImpairment (lynxRow: LynxRow) : HearingImpairment =
        match lynxRow.intake_hearing_loss with
        | true -> Yes
        | false -> No

    let getMobilityImpairment (lynxRow: LynxRow) : MobilityImpairment =
        match lynxRow.intake_mobility with
        | true -> Yes
        | false -> No

    let getCommunicationImpairment (lynxRow: LynxRow) : CommunicationImpairment =
        match lynxRow.intake_communication with
        | true -> Yes
        | false -> No

    let getCognitiveImpairment (lynxRow: LynxRow) : CognitiveImpairment =
        [ lynxRow.intake_alzheimers
        ; lynxRow.intake_learning_disability
        ; lynxRow.intake_memory_loss
        ]
        |> hasImpairment

    let getMentalHealthImpairment (lynxRow: LynxRow) : MentalHealthImpairment =
        [ (lynxRow.intake_mental_health |> optstringToBool)
        ; lynxRow.intake_substance_abuse
        ]
        |> hasImpairment

    let getOtherImpairment (lynxRow: LynxRow) : OtherImpairment =
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
        ; (lynxRow.intake_allergies |> optstringToBool)
        ; lynxRow.intake_dexterity
        ]
        |> hasImpairment

    AgeRelatedImpairments
        ( getHearingImpairment lynxRow
        , getMobilityImpairment lynxRow
        , getCommunicationImpairment lynxRow
        , getCognitiveImpairment lynxRow
        , getMentalHealthImpairment lynxRow
        , getOtherImpairment lynxRow
        )

let getTypeOfResidence (lynxRow: LynxRow) : TypeOfResidence =
    match lynxRow.intake_residence_type with
    | Some s when s = (toOIBString PrivateResidence) ->
        PrivateResidence
    | Some s when s = (toOIBString SeniorIndependentLiving) ->
        SeniorIndependentLiving
    | Some s when s = (toOIBString AssistedLivingFacility) ->
        TypeOfResidence.AssistedLivingFacility
    | Some s when s = (toOIBString NursingHome) ->
        TypeOfResidence.NursingHome
    | Some s when s = (toOIBString Homeless) ->
        Homeless
    // Historical LYNX options
    | Some "Community Residential" ->
        SeniorIndependentLiving
    | Some "Skilled Nursing Care" ->
        TypeOfResidence.NursingHome
    | Some "Assisted Living" ->
        TypeOfResidence.AssistedLivingFacility
    | Some _
    | None ->
        failwith $"Type of residence {lynxRow.intake_residence_type} not in OIB report. Client: {lynxRow.contact_id} {getClientName lynxRow}."

let getSourceOfReferral (lynxRow: LynxRow) : SourceOfReferral =

let createDemographicsRow (lynxRow: LynxRow) (grantYearStart: System.DateOnly) =
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
