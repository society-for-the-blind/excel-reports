module ExcelReports.OIBTypes

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

type IOIBString =
    abstract member ToOIBString: unit -> string

type YesOrNo =
    | Yes
    | No

    // Implement IOIBString interface for yesOrNo
    interface IOIBString with
        member this.ToOIBString() =
            match this with
            | Yes -> "Yes"
            | No  -> "No"

let toOIBString (value: IOIBString) =
    value.ToOIBString ()

type ClientName =
| ClientName of string

    interface IOIBString with
        member this.ToOIBString() =
            match this with
            | ClientName name -> name

// TODO 2023-12-04_2335
// Create interface and implement it for each type
// to convert to the proper value according the
// Excel OIB report.
// https://learn.microsoft.com/en-us/dotnet/fsharp/language-reference/discriminated-unions#members
type IndividualsServed =
    | NewCase
    | PriorCase

    interface IOIBString with
        member this.ToOIBString() =
            match this with
            | NewCase   -> "Case open between Oct. 1 - Sept. 30"
            | PriorCase -> "Case open prior to Oct. 1"
//   ("B7:D7",
//    [("B2", "Case open prior to Oct. 1");
//     ("C2", "Case open between Oct. 1 - Sept. 30")])

type AgeAtApplication =
    | From55To64
    | From65To74
    | From75To84
    | From85AndOlder

    interface IOIBString with
        member this.ToOIBString() =
            match this with
            | From55To64 -> "55-64"
            | From65To74 -> "65-74"
            | From75To84 -> "75-84"
            | From85AndOlder -> "85 and older"
//   ("E7:I7",
//    [("E2", "55-64"); ("F2", "65-74"); ("G2", "75-84"); ("H2", "85 and older")])

type Gender =
    | Female
    | Male
    | DidNotSelfIdentifyGender

    // Implement IOIBString interface for Gender
    interface IOIBString with
        member this.ToOIBString() =
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
    | DidNotSelfIdentifyEthnicity
    | TwoOrMoreRaces

    // Implement IOIBString interface for Race
    interface IOIBString with
        member this.ToOIBString() =
            match this with
            | NativeAmerican                  -> "Native American"
            | Asian                           -> "Asian"
            | AfricanAmerican                 -> "African American"
            | PacificIslanderOrNativeHawaiian -> "Pacific Islander or Native Hawaiian"
            | White                           -> "White"
            | DidNotSelfIdentifyEthnicity     -> "Did Not Self-Identify Ethnicity"
            | TwoOrMoreRaces                  -> "Two or More Races"
//   ("N7:U7",
//    [("N2", "American Indian or Alaska Native"); ("O2", "Asian");
//     ("P2", "Black or African American");
//     ("Q2", "Native Hawaiian or Pacific Islander"); ("R2", "White");
//     ("S2", "Did not self identify Race"); ("T2", "2 or More Races")])

// !!! TODO / PONDER !!! 2023-12-04_2332
// If this is true, then `TwoOrMoreRaces: Race` has to be set for the client.
// REASON:
// Because our current system treats ethnicity and
// race in one list... Therefore, it is would be
// wrong to represent this constraint in the type
// system and should be handled in the constructor
// of DemographicsRow (see TODO 2023-12-04_2337).
type HispanicOrLatino = YesOrNo
type Ethnicity = HispanicOrLatino
// "V7" [|"Yes"; "No"|]

type DegreeOfVisualImpairment =
    | TotallyBlind
    | LegallyBlind
    | SevereVisionImpairment

    // Implement IOIBString interface for DegreeOfVisualImpairment
    interface IOIBString with
        member this.ToOIBString() =
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

    // Implement IOIBString interface for MajorCauseOfVisualImpairment
    interface IOIBString with
        member this.ToOIBString() =
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

type AgeRelatedImpairments =
    ( HearingImpairment
    * MobilityImpairment
    * CommunicationImpairment
    * CognitiveImpairment
    * MentalHealthImpairment
    * OtherImpairment
    )

    // Implement IOIBString interface for AgeRelatedImpairments
    // interface IOIBString with
    //     member this.ToOIBString() =
    //         match this with
    //         | HearingImpairment       true
    //         | MobilityImpairment      true
    //         | CommunicationImpairment true
    //         | CognitiveImpairment     true
    //         | MentalHealthImpairment  true
    //         | OtherImpairment         true -> "Yes"
    //         |                            _ -> "No"
//  "AG7:AL7" "[|"Yes"; "No"|]

type TypeOfResidence =
    | PrivateResidence
    | SeniorIndependentLiving
    | AssistedLivingFacility
    | NursingHome
    | Homeless

    // Implement IOIBString interface for TypeOfResidence
    interface IOIBString with
        member this.ToOIBString() =
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

// Implement IOIBString interface for SourceOfReferral
    interface IOIBString with
        member this.ToOIBString() =
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

    interface IOIBString with
        member this.ToOIBString() =
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

// A row in the "PART III-DEMOGRAPHICS" sheet
type DemographicsRow =
    ( ClientName                   // "A7"
    * IndividualsServed            // "B7:D7"
    * AgeAtApplication             // "E7:I7"
    * Gender                       // "J7:M7"
    * Race                         // "N7:U7"
    // TODO 2023-12-04_2337
    // Add smart constructor for validation to cater for 2023-12-04_2337
    * Ethnicity                    // "V7"
    * DegreeOfVisualImpairment     // "W7:Z7"
    * MajorCauseOfVisualImpairment // "AA7:AF7"
    * AgeRelatedImpairments        // "AG7:AL7"
    * TypeOfResidence              // "AM7:AR7"
    * SourceOfReferral             // "AS7:BE7"
    * County                       // "BF7"
    )

// A. Clinical/functional Vision Assessments and Services
// Vision  Assessment (Screening/ Exam/evaluation	Surgical or Therapeutic Treatment
type VisionAssessment = YesOrNo
type SurgicalOrTherapeuticTreatment = YesOrNo

type ClinicalFunctionalVisionAssessmentsAndServices =
    ( VisionAssessment
    * SurgicalOrTherapeuticTreatment
    )

// B. Assistive Technology Devices and Services
// AT Goal Outcomes
// C. Independent Living and Adjustment Services
// IL/A Service Goal Outcomes
// D. Supportive Services
// Case Status
// Living Situation Outcomes
// Home and Community involvement Outcomes
// Employment Outcome
// Number of Services
// County

// A row in the "PART IV-V-SERVICES AND OUTCOMES" sheet
// type ServicesRow =
    // ( // "A7"  is  a formula  pulling  `ClientName`
      // from  the Demographics  sheet's "A7"  cell
      // (see `DemographicsRow`)


// NOTE 2023-12-04_2257
// Not sure if a type unifying row types will be
// needed, so just in case.
// type ReportRows