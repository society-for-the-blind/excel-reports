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

type IndividualsServed =
| NewCase
| PriorCase

type AgeAtApplication =
| From55To64
| From65To74
| From75To84
| From85AndOlder

type Gender =
| Female
| Male
| DidNotSelfIdentifyGender
// [("J2", "Female"); ("K2", "Male"); ("L2", "Did Not Self-Identify Gender")])

type Race =
| NativeAmerican
| Asian
| AfricanAmerican
| PacificIslanderOrNativeHawaiian
| White
| DidNotSelfIdentifyEthnicity
| TwoOrMoreRaces

type DegreeOfVisualImpairment =
| TotallyBlind
| LegallyBlind
| SevereVisionImpairment

type MajorCauseOfVisualImpairment =
| MacularDegeneration
| DiabeticRetinopathy
| Glaucoma
| Cataracts
| OtherCausesOfVisualImpairment

type Demographics =
| IndividualsServed of IndividualsServed
| AgeAtApplication of AgeAtApplication
| Gender of Gender
| Race of Race
// !!! TODO / PONDER !!!
// If this is true, then Demographics (Race TwoOrMoreRaces) has to be set for the client.
// This does not make much sense but this is how it is done.
| Ethnicity of bool
| DegreeOfVisualImpairment of DegreeOfVisualImpairment
| MajorCauseOfVisualImpairment of MajorCauseOfVisualImpairment

// H. Other Age-Related Impairments
//   Hearing Impairment
//   Mobility Impairment
//   Communication Impairment
//   Cognitive or Intellectual Impairment
//   Mental Health Impairment
//   Other Impairment
// |"Yes"; "No"|

// I. Type of Residence
//  [("AM2", "Private Residence"); ("AN2", "Senior Independent Living");
//     ("AO2", "Assisted Living Facility"); ("AP2", "Nursing Home");
//     ("AQ2", "Homeless")])

// J. Source of Referral
//  [("AS2", "Eye Care Provider"); ("AT2", "Physician/ Medical Provider");
//     ("AU2", "State VR Agency"); ("AV2", "Social Service");
//     ("AW2", "Veterans Administration"); ("AX2", "Senior Program");
//     ("AY2", "Assisted Living Facility"); ("AZ2", "Nursing Home");
//     ("BA2", "Independent Living Center"); ("BB2", "Family or Friend");
//     ("BC2", "Self-Referral"); ("BD2", "Other")])

// County
//  ["Alameda"; "Alpine"; "Amador"; "Butte"; "Calaveras"; "Colusa";
//    "Contra Costa"; "Del Norte"; "El Dorado"; "Fresno"; "Glenn"; "Humboldt";
//    "Imperial"; "Inyo"; "Kern"; "Kings"; "Lake"; "Lassen"; "Los Angeles";
//    "Madera"; "Marin"; "Mariposa"; "Mendocino"; "Merced"; "Modoc"; "Mono";
//    "Monterey"; "Napa"; "Nevada"; "Orange"; "Placer"; "Plumas"; "Riverside";
//    "Sacramento"; "San Benito"; "San Bernardino"; "San Diego"; "San Francisco";
//    "San Joaquin"; "San Luis Obispo"; "San Mateo"; "Santa Barbara";
//    "Santa Clara"; "Santa Cruz"; "Shasta"; "Sierra"; "Siskiyou"; "Solano";
//    "Sonoma"; "Stanislaus"; "Sutter"; "Tehama"; "Trinity"; "Tulare"; "Tuolumne";
//    "Ventura"; "Yolo"; "Yuba"]

type ListValidationValues =
| Demographics of Demographics

// TODO create the list validation types and then collect them in a union type, then replace 'a with it
//      Would this work?...
let listValidationTypeToString (validationValue: ListValidationValues) =
    match validationValue with
    | Demographics (IndividualsServed PriorCase) -> "Case open prior to Oct. 1"
    | Demographics (IndividualsServed NewCase)   -> "Case open between Oct. 1 - Sept. 30"
    | Demographics (AgeAtApplication From55To64)     -> "55-64"
    | Demographics (AgeAtApplication From65To74)     -> "65-74"
    | Demographics (AgeAtApplication From75To84)     -> "75-84"
    | Demographics (AgeAtApplication From85AndOlder) -> "85 and older"
    | Demographics (Gender Female)                   -> "Female"
    | Demographics (Gender Male)                     -> "Male"
    | Demographics (Gender DidNotSelfIdentifyGender) -> "Did Not Self-Identify Gender"
    | Demographics (Race NativeAmerican)                  -> "American Indian or Alaska Native"
    | Demographics (Race Asian)                           -> "Asian"
    | Demographics (Race AfricanAmerican)                 -> "Black or African American"
    | Demographics (Race PacificIslanderOrNativeHawaiian) -> "Native Hawaiian or Pacific Islander"
    | Demographics (Race White)                           -> "White"
    | Demographics (Race DidNotSelfIdentifyEthnicity)     -> "Did not self identify Race"
    | Demographics (Race TwoOrMoreRaces)                  -> "2 or More Races"
    | Demographics (Ethnicity true)  -> "Yes"
    | Demographics (Ethnicity false) -> "No"
    | Demographics (DegreeOfVisualImpairment TotallyBlind)           -> "Totally Blind"
    | Demographics (DegreeOfVisualImpairment LegallyBlind)           -> "Legally Blind"
    | Demographics (DegreeOfVisualImpairment SevereVisionImpairment)            -> "Severe Vision Impairment"
    | Demographics (MajorCauseOfVisualImpairment MacularDegeneration)           -> "Macular Degeneration"
    | Demographics (MajorCauseOfVisualImpairment DiabeticRetinopathy)           -> "Diabetic Retinopathy"
    | Demographics (MajorCauseOfVisualImpairment Glaucoma)                      -> "Glaucoma"
    | Demographics (MajorCauseOfVisualImpairment Cataracts)                     -> "Cataracts"
    | Demographics (MajorCauseOfVisualImpairment OtherCausesOfVisualImpairment) -> "Other causes of visual impairment"

// To get list validation values quickly
let gv (cell: ICell) =
    (findDataValidation cell |> Option.get).ValidationConstraint.Formula1
    |> fun str -> str.Replace("$","")
    |> convertCellRangeToList
    |> List.map (fun address -> (address, (getCell (cell.Sheet.Workbook :?> XSSFWorkbook) 3 address).StringCellValue) )
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
