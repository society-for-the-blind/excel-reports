module ExcelReports.LynxData

(*
#r "nuget: Npgsql.FSharp, 5.7.0";;
#load "ExcelReports/LynxData.fs";;
open ExcelReports.LynxData;;
*)

open System
open System.Reflection
open Npgsql.FSharp
open System.Text.RegularExpressions

open OIBTypes

type LynxColumn = {
    ColumnName : string;
    ColumnType : Type;
}

let flip (f: 'a -> 'b -> 'c) (x: 'b) (y: 'a) = f y x

// type RowReaderMember =
// | Int of int
// | Text of string
// | Bool of bool
// | DateOnly of System.DateOnly
// | oextOrNone of string option

let (|GetType|_|) (ct: Type) (t: Type) =
    match Regex("Option").Match(t.FullName).Success with
    | true ->
        Regex("\[(System.*?),").Match(t.FullName).Groups.[1].Value
        |> fun (extractedType: string) -> (extractedType = ct.FullName)
        |> fun (isMatch: bool) -> ( isMatch, "option" )
        |> Some
    | false ->
        Some <| ( (t.FullName = ct.FullName), "" )

// NOTE How to add new F# -> Npgsql.FSharp mapping?
//      Check the `RowReader` type's members.
let typeToRowReaderMember (t: Type) =
    let intType      = typeof<int>
    let stringType   = typeof<string>
    let boolType     = typeof<bool>
    let dateOnlyType = typeof<System.DateOnly>
    let dateTimeType = typeof<System.DateTime>

    match t with
    | GetType intType      (true, "") -> "int"
    | GetType stringType   (true, "") -> "text"
    | GetType boolType     (true, "") -> "bool"
    | GetType dateOnlyType (true, "") -> "dateOnly"
    | GetType dateTimeType (true, "") -> "dateTime"
    | GetType dateTimeType (true, "option") -> "dateTimeOrNone"
    | GetType dateOnlyType (true, "option") -> "dateOnlyOrNone"
    | GetType stringType   (true, "option") -> "textOrNone"
    | GetType boolType     (true, "option") -> "boolOrNone"
    | _ -> failwith $"TYPE {t.FullName} NOT IMPLEMENTED in `LynxData` MODULE'S `typeToRowReaderMember`."

let toLynxColumn (fieldInfo: FieldInfo) : LynxColumn =
    // NOTE 2023-12-01_2227
    // The last character of the field name was "@" in all my
    // experiments, so decided to get sloppy and just remove the last
    // char instead of checking for it.
    let delLastChar (str: string) : string =
        str.Substring(0, str.Length - 1)

    {
        ColumnName = fieldInfo.Name |> delLastChar
        ColumnType = fieldInfo.FieldType;
    }

let deleteUpToFirstUnderscore (str: string) =
    let index = str.IndexOf('_')
    if index >= 0 then
        str.Substring(index + 1)
    else
        str

// let fieldNamesAndTypes = getRecordFieldNamesAndTypes<LynxRow>()

// TODO 2023-12-01_1342
//      Remove the hard-coded password.
let connectionString = "postgres://postgres:XntSrCoEEZtiacZrx2m7jR5htEoEfYyoKncfhNmnPrLqPzxXTU5nxM@192.168.64.4:5432/lynx"

// NOTE 2023-12-14_1221 Why are most record fields options?
// Because  the LYNX  database  tables  have almost  no
// constraints (practically used as an Excel workbook),
// so it is  just safer to treat any  returned value as
// invalid.
type LynxRow = {

    // NOTE Naming convention of the record fields
    //
    //      <table_alias>_<column_name>
    //
    //      where `<table_alias>`-es are defined
    //      in the  `lynxQuery` function in the
    //      `joins` variable.

    // NOTE How to use SQL aliases?
    //
    //      See NOTE "add_SQL_aliases" before `lynxQuery`
    //      function.

    contact_id          :    int; // aliased! see NOTE "add_SQL_aliases"
    contact_last_name   : string option;
    contact_first_name  : string option;
    contact_middle_name : string option;

    intake_id          : int; // aliased! see NOTE "add_SQL_aliases"
    intake_intake_date : System.DateOnly option;
    intake_birth_date  : System.DateOnly option;
    intake_gender      :          string option;

    // TODO 2023-12-02_2230
    // This one belongs to 2 OIB demographics columns (race, ethnicity).
    // The race is a dropdown with pre-defined values, ethnicity only
    // means whether the person is Hispanic or not. If Hispanic, then
    // the "ethnicity" column will say "yes" and the "race" column will
    // ALWAYS be "2 or More Races".
    intake_ethnicity       : string option; // race
    intake_other_ethnicity : string option  // ethnicity (i.e., Hispanic or not)
    intake_degree          : string option; // degree of visual impairment
    intake_eye_condition   : string option; // major cause of visual impairment

    // H. Other Age-Related Impairments
    intake_hearing_loss  : bool option; // hearing impairment
    intake_mobility      : bool option; // mobility impairment
    intake_communication : bool option; // communication impairment

    // NOTE 2023-12-10_2226
    // `lynx_intake`  table's `geriatric`  column seems  to
    // map directly to the  OIB report's "Other Impairment"
    // column   in  the   "Demographics"   sheet,  but   we
    // decided  to also  "OR" it  with the  other remainder
    // health-related  columns that  don't belong  anywhere
    // else.
    intake_geriatric           :   bool option; // | other impairment
    intake_stroke              :   bool option; // |
    intake_seizure             :   bool option; // |
    intake_migraine            :   bool option; // |
    intake_heart               :   bool option; // |
    intake_diabetes            :   bool option; // |
    intake_dialysis            :   bool option; // |
    intake_cancer              :   bool option; // |
    intake_arthritis           :   bool option; // |
    intake_high_bp             :   bool option; // |
    intake_neuropathy          :   bool option; // |
    intake_pain                :   bool option; // |
    intake_asthma              :   bool option; // |
    intake_musculoskeletal     :   bool option; // |
    intake_allergies           : string option; // |
    intake_dexterity           :   bool option; // |

    intake_alzheimers          : bool option; // |
    intake_memory_loss         : bool option; // | cognitive impairment
    intake_learning_disability : bool option; // |

    intake_mental_health       : string option; // | mental health impairment
    intake_substance_abuse     :   bool option;  // |

    intake_residence_type      : string option; // | // type of residence
    intake_referred_by         : string option; // | // source of referral

    // Without adding  `note_id` only the joint  rows would
    // show  up where  the field  values are  distinct. Not
    // sure what  the benefit  of forcing  ALL notes  to be
    // present, but being explicit feels better.
    note_id             : int; // aliased! see NOTE "add_SQL_aliases"
    note_at_devices     :            bool option;
    note_at_services    :            bool option;
    note_orientation    :            bool option;
    note_dls            :            bool option;
    note_communications :            bool option;
    note_advocacy       :            bool option;
    note_counseling     :            bool option;
    note_information    :            bool option;
    note_support        :            bool option;
    note_services       :            bool option;
    note_note_date      : System.DateOnly option;

    plan_id                      :    int; // aliased! see NOTE "add_SQL_aliases"
    plan_plan_name               : string option;
    plan_at_outcomes             : string option;
    plan_community_plan_progress : string option;
    plan_ila_outcomes            : string option;
    plan_living_plan_progress    : string option;
    plan_plan_date               : System.DateOnly option;
    plan_modified                : System.DateTime option;

    mostRecentAddress_id       :    int; // aliased! see NOTE "add_SQL_aliases"
    mostRecentAddress_modified : System.DateTime option;
    mostRecentAddress_county   : string option;
}

// see TODO 2024-02-19_1348 rename_lynx_prefix_oib_report
type LynxQuery = LynxRow list

type LynxData = {
    grantYearStart : System.DateOnly;
    grantYearEnd   : System.DateOnly;
    lynxQuery : LynxQuery;
}

// Cannot be moved inside a function because of FS0665.
let getRecordFieldNamesAndTypes<'T, 'U> (mapper: FieldInfo -> 'U) =
    typeof<'T>.GetFields(BindingFlags.Public ||| BindingFlags.Instance)
    |> Array.map mapper

type Quarter =
    | Q1
    | Q2
    | Q3
    | Q4

// TODO 2024-02-19_1348 rename_lynx_prefix_oib_report
//
// The  `lynxQuery` function  and the  types `LynxRow`,
// `LynxColumn`,   `LynxData`  (and   probably  others)
// should  be renamed  to reflect  that they  relate to
// the  quarterly OIB  report (e.g.,  `oibReportQuery`,
// `OIBReportRow`,  `OIBReportColumn`,  `OIBReportData`
// respectively).

// NOTE: add_SQL_aliases
// 1. Update `match` in the `queryColumns` variable.
// 2. Update `function` clause in the `exeReader` function.
// 3. If something is still amiss, uncommend the `printfn` statements.
let lynxQuery (connectionString: string) (quarter: Quarter) (grantYear: int): LynxData =

    let toStartEndDates (q: Quarter) (grantYear: int) =
        let startDate =
            match q with
            | Q1 -> System.DateOnly (grantYear    , 10, 1)
            | Q2 -> System.DateOnly (grantYear + 1,  1, 1)
            | Q3 -> System.DateOnly (grantYear + 1,  4, 1)
            | Q4 -> System.DateOnly (grantYear + 1,  7, 1)

        ( startDate
        , startDate.AddMonths(3).AddDays(-1)
        )

    let grantYearStart = toStartEndDates Q1 grantYear |> fst
    let grantYearEnd   = toStartEndDates Q4 grantYear |> snd

    // `quarterStart` will always be `grantYearStart`
    // because the OIB report is a cumulative one
    let quarterEnd = toStartEndDates quarter grantYear |> snd

    let (lynxCols: LynxColumn array) =
        getRecordFieldNamesAndTypes<LynxRow,LynxColumn> toLynxColumn

    let replaceFirstOccurrence (str: string) (oldValue: char, newValue: char) =
        let index = str.IndexOf(oldValue)
        if index >= 0 then
            str.Remove(index, 1).Insert(index, newValue.ToString())
        else
            str

    // SELECT columns generated from LynxRow type.
    let queryColumns =
        lynxCols
        |> Array.map (fun { ColumnName = n; ColumnType = _ } ->
            let alias =
                match n with
                | "contact_id" as a -> a + " AS contact_id"
             // | "address_id" as a -> a + " AS address_id"
                | "intake_id"  as a -> a + " AS intake_id"
                | "note_id"    as a -> a + " AS note_id"
                | "plan_id"    as a -> a + " AS plan_id"
                | "plan_modified" as a -> a + " AS plan_modified"
                | rest -> rest
            replaceFirstOccurrence alias ('_', '.');)
        |> String.concat ", "

    let joins = """
         lynx_sipnote AS note
    JOIN lynx_contact AS contact ON    note.contact_id = contact.id
    JOIN lynx_sipplan AS plan    ON   note.sip_plan_id = plan.id
    JOIN lynx_intake  AS intake  ON  intake.contact_id = contact.id
    JOIN (
        SELECT address.id, address.contact_id, address.modified, address.county
        FROM lynx_address AS address
        JOIN (
            SELECT contact_id, MAX(modified) AS most_recent
            FROM lynx_address
            GROUP BY contact_id
        ) AS subq ON address.contact_id = subq.contact_id AND address.modified = subq.most_recent
    ) AS mostRecentAddress ON mostRecentAddress.contact_id = contact.id
    """

    let baseSelect = "SELECT " + queryColumns + " FROM " + joins

    let whereClause = $"WHERE note.note_date >= '{grantYearStart.ToString()}'::date AND note.note_date < '{quarterEnd.ToString()}'::date"

    // NOTE 2023-12-01_1347 Should be irrelevant.
    // let groupByClause = "GROUP BY " + queryColumns
    // let orderByClause = "ORDER BY CONCAT(c.last_name, ', ', c.first_name)"

    let query = $"{baseSelect} {whereClause}" // + "{groupByClause} {orderByClause}"

    printfn $"QUERY: {query}"

    let exeReader (read: RowReader) : LynxRow =

        let callMethodDynamically (instance: obj) (methodName: string) (args: obj[]) =
            let methodInfo = instance.GetType().GetMethod(methodName)
            methodInfo.Invoke(instance, args)

        let lynxRowType = typeof<LynxRow>
        let constructor = lynxRowType.GetConstructors().[0]
        let constructorArgs =
            lynxCols
            |> Array.map (fun {ColumnName = n; ColumnType = t} ->
                // printfn $"COLUMN NAME: {n}"
                n
                |> function
                    | "contact_id" as a -> a
                 // | "address_id" as a -> a
                    | "intake_id"  as a -> a
                    | "note_id"    as a -> a
                    | "plan_id"    as a -> a
                    | "plan_modified"    as a -> a
                    | rest -> deleteUpToFirstUnderscore rest
                |> fun columnName ->
                    // printfn $"COLUMN NAME: {columnName}"
                    [| box columnName |]
                |> callMethodDynamically read (typeToRowReaderMember t)
                |> box
            )

        constructor.Invoke(constructorArgs) :?> LynxRow

        // {
        //     ContactID = read.int "id";
        //     LastName = read.text "last_name";
        //     ...
        // }

    connectionString
    |> Sql.connect
    |> Sql.query query
    |> Sql.execute exeReader
    |> fun (q: LynxQuery) ->
        { grantYearStart = grantYearStart
        // Only used in `getAgeAtApplication` in `Library.fs`
        ; grantYearEnd = grantYearEnd
        ; lynxQuery = q
        }

// === SqlHydra EXPERIMENTS ===
// User ID=postgres;Password=XntSrCoEEZtiacZrx2m7jR5htEoEfYyoKncfhNmnPrLqPzxXTU5nxM;Host=192.168.64.4;Port=5432;Database=lynx;

// $ dotnet fsi
// Microsoft (R) F# Interactive version 12.8.0.0 for F# 8.0
// Copyright (c) Microsoft Corporation. All Rights Reserved.

// For help type #help;;

// >  #r "nuget: Npgsql.FSharp, 5.7.0";;
// [Loading /Users/toraritte/.packagemanagement/nuget/Cache/697d8ca5b71fe39e0b2bf72bb58c700d58b82d6d086bcfc1fa356cce2708e407.fsx]
// module FSI_0003.
//        697d8ca5b71fe39e0b2bf72bb58c700d58b82d6d086bcfc1fa356cce2708e407

// > #r "nuget: SqlHydra.Query, 2.2.1";;
// [Loading /Users/toraritte/.packagemanagement/nuget/Cache/177be160dcb44a4a927d2619eda16eee3526dd5810c7ef37a6d4f9fd4544ce0d.fsx]
// module FSI_0002.
//        177be160dcb44a4a927d2619eda16eee3526dd5810c7ef37a6d4f9fd4544ce0d

// > #r "nuget: SqlHydra.Cli, 2.3.0";;

// /Users/toraritte/dev/clones/dotNET/slate-excel-reports/stdin(1,1): error FS0999: /Users/toraritte/.packagemanagement/nuget/Projects/85296--b0cee205-014b-423f-951e-e8bd674cb3f1/Proje
// ct.fsproj : error NU1202: Package SqlHydra.Cli 2.3.0 is not compatible with net8.0 (.NETCoreApp,Version=v8.0). Package SqlHydra.Cli 2.3.0 supports:

// > #r "nuget: SqlHydra.Cli, 2.3.1";;

// /Users/toraritte/dev/clones/dotNET/slate-excel-reports/stdin(1,1): error FS0999: /Users/toraritte/.packagemanagement/nuget/Projects/85296--b0cee205-014b-423f-951e-e8bd674cb3f1/Proje
// ct.fsproj : error NU1202: Package SqlHydra.Cli 2.3.1 is not compatible with net8.0 (.NETCoreApp,Version=v8.0). Package SqlHydra.Cli 2.3.1 supports:
// ====================

let pairwiseFold (f: 'a -> 'a -> 'b) (xs: 'a seq) =
    xs
    |> Seq.pairwise
    |> Seq.map ( fun (pair) -> pair ||> f )
    |> Seq.distinct

open System.Reflection

let compareRecords record1 record2 =
    let recordType = record1.GetType()
    let fields = recordType.GetFields(BindingFlags.Public ||| BindingFlags.Instance)

    fields
    |> Array.filter (fun field -> not (obj.Equals(field.GetValue(record1), field.GetValue(record2))))
    |> Array.map (fun field -> (field.Name, field.GetValue(record1), field.GetValue(record2)))

// let filterByDifferingAddressCounty (lynxQuery: LynxQuery) =
//     lynxQuery
//     |> Seq.groupBy (fun row -> row.contact_id)
//     |> Seq.filter (
//             fun (_, rows) ->
//                 rows
//                 |> Seq.map (fun r -> r.address_id)
//                 |> Seq.distinct
//                 |> Seq.length > 1
//             )
//     |> Seq.map (
//             fun (cid, rows) ->
//                 let compared =
//                     pairwiseFold
//                         compareRecords
//                         rows
//                 ( cid, compared )
//             )

// let aaa record1 record2 =
//     match (compareRecords record1 record2) with
//     | [|("address_id@", id1, id2)|] -> Seq.maxBy (fun r -> r.address_id) [record1; record2]

// Seq.map (fun r -> (r.address_id, r.contact_last_name));;

//     lynxQuery
//     |> Seq.filter (fun row -> not (List.contains row filteredRows))
//     |> Seq.toList


// TO FIND CLIENTS WHO HAVE MULTIPLE COUNTIES IN THEIR ADDRESS (because another crucial constraint is missing from LYNX)
// qq.lynxQuery |> List.map (flip createDemographicsRow <| q.grantYearStart) |> Seq.groupBy (function | ((_, Ok c) :: _) -> toOIBString c | ((_, Error e) :: _) -> e) |> Seq.filter (fun (_, r) -> r |> Seq.distinct |> Seq.length > 1 ) |> Seq.map (fun (n, r) -> (n, (r |> Seq.distinct |> Seq.toList)));;

