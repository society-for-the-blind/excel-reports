module ExcelReports.OIBTypes

(*
#load "ExcelReports/OIBTypes.fs";;
open ExcelReports.OIBTypes;;
*)

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

type IStringable =
    abstract member Stringify: unit -> string

type YesOrNo =
    | Yes
    | No

    // Implement IStringable interface for yesOrNo
    interface IStringable with
        member this.Stringify() =
            match this with
            | Yes -> "Yes"
            | No  -> "No"

let stringify (value: IStringable) =
    value.Stringify ()

type ClientName =
| ClientName of string

    interface IStringable with
        member this.Stringify() =
            match this with
            | ClientName name -> name

// The name is misleading (just look at the union
// cases), but this is what the OIB report calls it.
type IndividualsServed =
    | NewCase
    | PriorCase

    interface IStringable with
        member this.Stringify() =
            match this with
            | NewCase   -> "Case open between Oct. 1 - Sept. 30"
            | PriorCase -> "Case open prior to Oct. 1"
//   ("B7:D7",
//    [("B2", "Case open prior to Oct. 1");
//     ("C2", "Case open between Oct. 1 - Sept. 30")])

type AgeAtApplication =
    | AgeBracket18To24
    | AgeBracket25To34
    | AgeBracket35To44
    | AgeBracket45To54
    | AgeBracket55To64
    | AgeBracket65To74
    | AgeBracket75To84
    | AgeBracket85AndOlder

    interface IStringable with
        member this.Stringify() =
            match this with
            | AgeBracket18To24 -> "18-24"
            | AgeBracket25To34 -> "25-34"
            | AgeBracket35To44 -> "35-44"
            | AgeBracket45To54 -> "45-54"
            | AgeBracket55To64 -> "55-64"
            | AgeBracket65To74 -> "65-74"
            | AgeBracket75To84 -> "75-84"
            | AgeBracket85AndOlder -> "85 and older"
//   ("E7:I7",
//    [("E2", "55-64"); ("F2", "65-74"); ("G2", "75-84"); ("H2", "85 and older")])

type Gender =
    | Female
    | Male
    | DidNotSelfIdentifyGender

    // Implement IStringable interface for Gender
    interface IStringable with
        member this.Stringify() =
            match this with
            | Female                   -> "Female"
            | Male                     -> "Male"
            | DidNotSelfIdentifyGender -> "Did Not Self-Identify Gender"
//   ("J7:M7",
//    [("J2", "Female"); ("K2", "Male"); ("L2", "Did Not Self-Identify Gender")])

type Race =
    | NativeAmerican
    | Asian
    | AfricanAmerican
    | PacificIslanderOrNativeHawaiian
    | White
    | DidNotSelfIdentifyRace
    | TwoOrMoreRaces

    // Implement IStringable interface for Race
    interface IStringable with
        member this.Stringify() =
            match this with
            | NativeAmerican                  -> "American Indian or Alaska Native"
            | Asian                           -> "Asian"
            | AfricanAmerican                 -> "Black or African American"
            | PacificIslanderOrNativeHawaiian -> "Native Hawaiian or Pacific Islander"
            | White                           -> "White"
            | DidNotSelfIdentifyRace          -> "Did not self identify Race"
            | TwoOrMoreRaces                  -> "2 or More Races"
//   ("N7:U7",
//    [("N2", "American Indian or Alaska Native"); ("O2", "Asian");
//     ("P2", "Black or African American");
//     ("Q2", "Native Hawaiian or Pacific Islander"); ("R2", "White");
//     ("S2", "Did not self identify Race"); ("T2", "2 or More Races")])

type HispanicOrLatino = YesOrNo
type Ethnicity = HispanicOrLatino
// "V7" [|"Yes"; "No"|]

type DegreeOfVisualImpairment =
    | TotallyBlind
    | LegallyBlind
    | SevereVisionImpairment

    // Implement IStringable interface for DegreeOfVisualImpairment
    interface IStringable with
        member this.Stringify() =
            match this with
            | TotallyBlind -> "Totally Blind"
            | LegallyBlind -> "Legally Blind"
            | SevereVisionImpairment -> "Severe Vision Impairment"
//   ("W7:Z7",
//    [("W2", "Totally Blind"); ("X2", "Legally Blind");
//     ("Y2", "Severe Vision Impairment")])

type MajorCauseOfVisualImpairment =
    | MacularDegeneration
    | DiabeticRetinopathy
    | Glaucoma
    | Cataracts
    | OtherCausesOfVisualImpairment

    // Implement IStringable interface for MajorCauseOfVisualImpairment
    interface IStringable with
        member this.Stringify() =
            match this with
            | MacularDegeneration           -> "Macular Degeneration"
            | DiabeticRetinopathy           -> "Diabetic Retinopathy"
            | Glaucoma                      -> "Glaucoma"
            | Cataracts                     -> "Cataracts"
            | OtherCausesOfVisualImpairment -> "Other Causes of Visual Impairment"
//   ("AA7:AF7",
//    [("AA2", "Macular Degeneration"); ("AB2", "Diabetic Retinopathy");
//     ("AC2", "Glaucoma"); ("AD2", "Cataracts");
//     ("AE2", "Other causes of visual impairment")])

type HearingImpairment       = YesOrNo
type MobilityImpairment      = YesOrNo
type CommunicationImpairment = YesOrNo
type CognitiveImpairment     = YesOrNo
type MentalHealthImpairment  = YesOrNo
type OtherImpairment         = YesOrNo

// TODO This may be an unnecessary complication.
// Yes, these column are grouped under this name,
// but it will just complicate the `DemographicsRow`
// type, which would otherwise be a tagged tuple.
type AgeRelatedImpairmentSubrow =
    AgeRelatedImpairments of
        ( HearingImpairment
        * MobilityImpairment
        * CommunicationImpairment
        * CognitiveImpairment
        * MentalHealthImpairment
        * OtherImpairment
        )
//  "AG7:AL7" "[|"Yes"; "No"|]

type TypeOfResidence =
    | PrivateResidence
    | SeniorIndependentLiving
    | AssistedLivingFacility
    | NursingHome
    | Homeless

    // Implement IStringable interface for TypeOfResidence
    interface IStringable with
        member this.Stringify() =
            match this with
            | PrivateResidence          -> "Private Residence"
            | SeniorIndependentLiving   -> "Senior Independent Living"
            | AssistedLivingFacility    -> "Assisted Living Facility"
            | NursingHome               -> "Nursing Home"
            | Homeless                  -> "Homeless"
// I. Type of Residence
//   ("AM7:AR7",
//    [("AM2", "Private Residence"); ("AN2", "Senior Independent Living");
//     ("AO2", "Assisted Living Facility"); ("AP2", "Nursing Home");
//     ("AQ2", "Homeless")])

type SourceOfReferral =
    | EyeCareProvider
    | PhysicianMedicalProvider
    | StateVRAgency
    | SocialService
    | VeteransAdministration
    | SeniorProgram
    | AssistedLivingFacility
    | NursingHome
    | IndependentLivingCenter
    | FamilyOrFriend
    | SelfReferral
    | Other

// Implement IStringable interface for SourceOfReferral
    interface IStringable with
        member this.Stringify() =
            match this with
            | EyeCareProvider           -> "Eye Care Provider"
            | PhysicianMedicalProvider  -> "Physician/ Medical Provider"
            | StateVRAgency             -> "State VR Agency"
            | SocialService             -> "Social Service"
            | VeteransAdministration    -> "Veterans Administration"
            | SeniorProgram             -> "Senior Program"
            | AssistedLivingFacility    -> "Assisted Living Facility"
            | NursingHome               -> "Nursing Home"
            | IndependentLivingCenter   -> "Independent Living Center"
            | FamilyOrFriend            -> "Family or Friend"
            | SelfReferral              -> "Self-Referral"
            | Other                     -> "Other"
// J. Source of Referral
//   ("AS7:BE7",
//    [("AS2", "Eye Care Provider"); ("AT2", "Physician/ Medical Provider");
//     ("AU2", "State VR Agency"); ("AV2", "Social Service");
//     ("AW2", "Veterans Administration"); ("AX2", "Senior Program");
//     ("AY2", "Assisted Living Facility"); ("AZ2", "Nursing Home");
//     ("BA2", "Independent Living Center"); ("BB2", "Family or Friend");
//     ("BC2", "Self-Referral"); ("BD2", "Other")])

type County =
    | Alameda       | Alpine       | Amador       | Butte      | Calaveras
    | Colusa        | ContraCosta  | DelNorte     | ElDorado   | Fresno
    | Glenn         | Humboldt     | Imperial     | Inyo       | Kern
    | Kings         | Lake         | Lassen       | LosAngeles | Madera
    | Marin         | Mariposa     | Mendocino    | Merced     | Modoc
    | Mono          | Monterey     | Napa         | Nevada     | Orange
    | Placer        | Plumas       | Riverside    | Sacramento | SanBenito
    | SanBernardino | SanDiego     | SanFrancisco | SanJoaquin | SanLuisObispo
    | SanMateo      | SantaBarbara | SantaClara   | SantaCruz  | Shasta
    | Sierra        | Siskiyou     | Solano       | Sonoma     | Stanislaus
    | Sutter        | Tehama       | Trinity      | Tulare     | Tuolumne
    | Ventura       | Yolo         | Yuba

    interface IStringable with
        member this.Stringify() =
            match this with
            | ContraCosta    -> "Contra Costa"
            | DelNorte       -> "Del Norte"
            | ElDorado       -> "El Dorado"
            | LosAngeles     -> "Los Angeles"
            | SanBenito      -> "San Benito"
            | SanBernardino  -> "San Bernardino"
            | SanDiego       -> "San Diego"
            | SanFrancisco   -> "San Francisco"
            | SanJoaquin     -> "San Joaquin"
            | SanLuisObispo  -> "San Luis Obispo"
            | SanMateo       -> "San Mateo"
            | SantaBarbara   -> "Santa Barbara"
            | SantaClara     -> "Santa Clara"
            | SantaCruz      -> "Santa Cruz"
            | _              -> string this
// "BF7"

// TODO 2023-12-18_2336
// Can this even be used in practice?
// type DemographicsColumn =
//     // Each union option is a column of singular or
//     // merged cells (see ranges in comments)
//     | A     of ClientName                   // "A7"
//     | B_D   of IndividualsServed            // "B7:D7"
//     | E_I   of AgeAtApplication             // "E7:I7"
//     | J_M   of Gender                       // "J7:M7"
//     | N_U   of Race                         // "N7:U7"
//     | V     of Ethnicity                    // "V7"
//     | W_Z   of DegreeOfVisualImpairment     // "W7:Z7"
//     | AA_AF of MajorCauseOfVisualImpairment // "AA7:AF7"
//     // Columns of `AgeRelatedImpairmentSubrow` "AG7:AL7"
//     | AG    of HearingImpairment
//     | AH    of MobilityImpairment
//     | AI    of CommunicationImpairment
//     | AJ    of CognitiveImpairment
//     | AK    of MentalHealthImpairment
//     | AL    of OtherImpairment
//     | AM_AR of TypeOfResidence              // "AM7:AR7"
//     | AS_BE of SourceOfReferral             // "AS7:BE7"
//     | BF    of County                       // "BF7"

// TODO 2023-12-20_2107 I don't think this is needed.
// A row in the "PART III-DEMOGRAPHICS" sheet
// type DemographicsRow =
//     // PONDER 2023-12-17_2233
//     // Should these be `DemoGraphicsColumn`? Or
//     // would that just be an exercise in pedantry?
//     // After all, the whole point of a tuple is that
//     // it can contain different types. Also, if
//     // `DemographicsColumn` is used, then it will be
//     // repeated 12 times ( e.g., `DemographicsColumn
//     // of DemographicsColumn * ...`), losing most of
//     // its meaning.
//     DemographicsRow of
//         ( ClientName                   // "A7"
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
//         // Why the `option` type? See TODO 2023-12-11_1617
//         // in `ExcelReports/LynxData.fs`.
//         * DegreeOfVisualImpairment     // "W7:Z7"
//         * MajorCauseOfVisualImpairment // "AA7:AF7"
//         // * AgeRelatedImpairmentSubrow   // "AG7:AL7"
//         * HearingImpairment // "AG7"
//         * MobilityImpairment // "AH7"
//         * CommunicationImpairment // "AI7"
//         * CognitiveImpairment // "AJ7"
//         * MentalHealthImpairment // "AK7"
//         * OtherImpairment // "AL7"
//         * TypeOfResidence              // "AM7:AR7"
//         * SourceOfReferral             // "AS7:BE7"
//         * County                       // "BF7"
//         )

// ---SERVICES----------------------------------------------------------

// type PlanDate =
//     | PlanDate of System.DateOnly

//     interface IStringable with
//         member this.Stringify() =
//             match this with
//             | PlanDate date -> date.ToString()

// type PlanId =
//     | PlanId of int

//     interface IStringable with
//         member this.Stringify() =
//             match this with
//             | PlanId id -> string(id)

type PlanModified =
    | PlanModified of System.DateTime

    interface IStringable with
        member this.Stringify() =
            match this with
            | PlanModified dt -> dt.ToString()

// A. Clinical/functional Vision Assessments and Services
type               VisionAssessment = YesOrNo // "B7"
type SurgicalOrTherapeuticTreatment = YesOrNo // "C7"

// B. Assistive Technology Devices and Services
type ReceivedAssistiveTechnologyServicesOrDevices = YesOrNo // "D7"

type IOIBOutcome =
    abstract member outcome : unit -> unit

type OutcomeA =
    | NotAssessed
    | AssessedWithImprovedIndependence
    | AssessedAndMaintainedIndependence
    | AssessedWithDecreasedIndependence

    interface IStringable with
        member this.Stringify() =
            match this with
            | NotAssessed                        -> "Not assessed"
            | AssessedWithImprovedIndependence   -> "Assessed with improved independence"
            | AssessedAndMaintainedIndependence  -> "Assessed and maintained independence"
            | AssessedWithDecreasedIndependence  -> "Assessed with decreased independence"

    interface IOIBOutcome with
        member this.outcome () = ()

type AssistiveTechnologyGoalOutcomes = OutcomeA //   ("E7:H7",
//   ("E7:H7",
//    [("E2", "Not assessed"); ("F2", "Assessed with improved independence");
//     ("G2", "Assessed and maintained independence");
//     ("H2", "Assessed with decreased independence")])

// type AssistiveTechnologyColumns =
//     // TODO 2023-12-07_1021
//     // If `ReceivedAssistiveTechnologyServicesOrDevices` is `No` then `AssistiveTechnologyGoalOutcomes` should be `NotAssessed`.
//     // !!!
//     // !!! Make it a generic constraint because this will be a theme in later columns.
//     // !!!
//     // PONDER Unlike 2023-12-04_2337 constraint,
//     // this one IS imposed by the OIB report. How
//     // to highlight this fact? (This is a value
//     // level constraint, so, as far as I know, it
//     // cannot be represented in the type system
//     // without dependent types.)
//     // PONDER ADDENDUM
//     // Tried to implement this constraint in the
//     // type system (see below) but then another
//     // "translation" type / function will be
//     // needed to match the report's layout.
//     //
//     //     // This illustrates the point above, but
//     //     // it's still bad: I would associate
//     //     // `ServicesDelivered` with a list of
//     //     // services and not with outcomes.
//     //     type WereAssistiveTechnologyServicesOrDevicesDelivered = // "D7"
//     //     | NoServicesOrDevicesDelivered
//     //     | ServicesDelivered of AssistiveTechnologyGoalOutcomes
//     //
//     // Suffice to say, the goal is to get the job done,
//     // and using `...Row` types is easier/quicker for now.
//     AssistiveTechnologyColumns of
//         ( ReceivedAssistiveTechnologyServicesOrDevices // "D7"
//         * AssistiveTechnologyGoalOutcomes              // "E7:H7"
//         )

// C. Independent Living and Adjustment Services
//
//    "Received IL/A Services" column is a formula computing its value from the following columns:
type ReceivedOrientationAndMobilityTraining = YesOrNo // "J7"
type ReceivedCommunicationSkills            = YesOrNo // "K7"
type ReceivedDailyLivingSkills              = YesOrNo // "L7"
type ReceivedAdvocacyTraining               = YesOrNo // "M7"
type ReceivedAdjustmentCounseling           = YesOrNo // "N7"
type ReceivedInformationAndReferral         = YesOrNo // "O7"
type ReceivedOtherServices                  = YesOrNo // "P7"

type IndependentLivingAndAdjustmentOutcomes = OutcomeA // "Q7:T7"

// D. Supportive Services
type ReceivedSupportiveService = YesOrNo // "U7"

type CaseStatus = // "V7"
    | Assessed
    | Pending

    interface IStringable with
        member this.Stringify() =
            match this with
            | Assessed   -> "Assessed"
            | Pending -> "Pending"

type OutcomeB =
    | Increased
    | Maintained
    | Decreased
    | NotAssessed

    interface IStringable with
        member this.Stringify() =
            match this with
            | Increased -> "Increased"
            | Maintained -> "Maintained"
            | Decreased -> "Decreased"
            | NotAssessed -> "Not Assessed"

    interface IOIBOutcome with
        member this.outcome () = ()

type             LivingSituationOutcomes = OutcomeB // "W7:Z7"
type HomeAndCommunityInvolvementOutcomes = OutcomeB // "AA7:AD7"

type EmploymentOutcomes = // "AE7:AH7"
    | NotInterested
    | LessLikely
    | Unsure
    | MoreLikely

    interface IStringable with
        member this.Stringify() =
            match this with
            | NotInterested -> "Not Interested"
            | LessLikely -> "Less Likely"
            | Unsure -> "Unsure"
            | MoreLikely -> "More Likely"

// A row in the "PART IV-V-SERVICES AND OUTCOMES" sheet
// type ServicesRow =
//     ServicesRow of
//         ( // "A7"  is  a formula  pulling  `ClientName`
//           // from  the Demographics  sheet's "A7"  cell
//           // (see `DemographicsRow`)
//           ClinicalFunctionalVisionAssessmentsAndServices // "B7:C7"
//           // TODO see 2023-12-07_1021 constraint
//         * AssistiveTechnologyColumns                // "D7:H7"
//         // "I7" is a formula calculating from "J7:P7" whether
//         // client received any IL/A services
//         )

// NOTE 2023-12-04_2257
// Not sure if a type unifying row types will be
// needed, so just in case.
// type ReportRows
