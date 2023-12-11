module ExcelReports.LynxData

// #r "nuget: SqlHydra.Query, 2.2.1";;
// #r "nuget: Npgsql.FSharp, 5.7.0";;
open System
open System.Reflection
open Npgsql.FSharp

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
// | TextOrNone of string option

let typeToRowReaderMember (t: Type) =
    match t with
    | _ when t = typeof<int> -> "int"
    | _ when t = typeof<string> -> "text"
    | _ when t = typeof<bool> -> "bool"
    | _ when t = typeof<System.DateOnly> -> "dateOnly"
    | _ when t.FullName.Contains("Option") &&
             t.FullName.Contains("String") -> "textOrNone"
    | _ -> failwith "NOT IMPLEMENTED: Type not supported."

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

// NOTE Naming convention of the record fields
//
//      <table_alias>_<column_name>
//
//      where `<table_alias>`-es are defined
//      in the  `lynxQuery` function in the
//      `joins` variable.
type LynxRow = {
    contact_id          : int;
    contact_last_name   : string;
    contact_first_name  : string;
    contact_middle_name : string option;

    intake_intake_date         : System.DateOnly;
    intake_birth_date          : System.DateOnly;
    intake_gender              : string option;
    // TODO 2023-12-02_2230
    // This one belongs to 2 OIB demographics columns (race, ethnicity).
    // The race is a dropdown with pre-defined values, ethnicity only
    // means whether the person is Hispanic or not. If Hispanic, then
    // the "ethnicity" column will say "yes" and the "race" column will
    // ALWAYS be "2 or More Races".
    intake_ethnicity           : string option; // race
    intake_other_ethnicity     : string option  // ethnicity (i.e., Hispanic or not)
    intake_degree              : string option; // degree of visual impairment
    intake_eye_condition       : string option; // major cause of visual impairment
    intake_hearing_loss        : bool;          // hearing impairment
    intake_mobility            : bool;          // mobility impairment
    intake_communication       : bool;          // communication impairment
    intake_stroke              : bool;          // |
    intake_seizure             : bool;          // |
    intake_alzheimers          : bool;          // | // cognitive or intellectual impairment
    intake_memory_loss         : bool;          // | // other impairment
    intake_learning_disability : bool;          // |
    intake_migraine            : bool;          // | TODO 2023-12-02_2157 talk to Shane how to untangle this
    intake_heart               : bool;          // |                      and what belongs to other impairments
    intake_diabetes            : bool;          // |
    intake_dialysis            : bool;          // |
    intake_cancer              : bool;          // |
    intake_arthritis           : bool;          // |
    intake_high_bp             : bool;          // |
    intake_neuropathy          : bool;          // |
    intake_pain                : bool;          // |
    intake_asthma              : bool;          // |
    intake_musculoskeletal     : bool;          // |
    intake_substance_abuse     : bool;          // |
    intake_allergies           : string option; // | !!!
    intake_geriatric           : bool;          // |
    intake_dexterity           : bool;          // |
    intake_residence_type      : string option; // | // type of residence
    intake_referred_by         : string option; // | // source of referral

    // TODO 2023-12-02_2153
    // This is a boolean in the OIB report, so I guess that when this is
    // not null, then it should be true.
    intake_mental_health       : string option // mental health impairment

    note_at_devices     : bool;
    note_orientation    : bool;
    note_dls            : bool;
    note_communications : bool;
    note_advocacy       : bool;
    note_counseling     : bool;
    note_information    : bool;
    note_support        : bool;
    note_note_date      : System.DateOnly;

    plan_plan_name               : string;
    plan_at_outcomes             : string option;
    plan_community_plan_progress : string option;
    plan_ila_outcomes            : string option;
    plan_living_plan_progress    : string option;
}

type LynxQuery = LynxRow list

type LynxData = {
    grantYearStart : System.DateOnly;
    lynxQuery : LynxQuery;
}

let getRecordFieldNamesAndTypes<'T, 'U> (mapper: FieldInfo -> 'U) =
    typeof<'T>.GetFields(BindingFlags.Public ||| BindingFlags.Instance)
    |> Array.map mapper

// let qtestodelete = connectionString |> Sql.connect |> Sql.query "select * from lynx_sipnote where id = 27555;" |> Sql.execute (fun (read: RowReader) -> read.text "note")

let lynxQuery (connectionString: string) (grantYear: int) : LynxData =

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
           replaceFirstOccurrence n ('_', '.');)
        |> String.concat ", "

    let joins = """
         lynx_sipnote AS note
    JOIN lynx_contact AS contact ON   note.contact_id = contact.id
    JOIN lynx_sipplan AS plan    ON  note.sip_plan_id = plan.id
    JOIN lynx_intake  AS intake  ON intake.contact_id = contact.id
    """

    let baseSelect = "SELECT " + queryColumns + " FROM " + joins

    let whereClause = $"WHERE note.note_date >= '{string grantYear}-10-01'::date AND note.note_date < '{string (grantYear+1)}-10-01'::date"

    // NOTE 2023-12-01_1347 Should be irrelevant.
    // let groupByClause = "GROUP BY " + queryColumns
    // let orderByClause = "ORDER BY CONCAT(c.last_name, ', ', c.first_name)"

    let query = $"{baseSelect} {whereClause}" // + "{groupByClause} {orderByClause}"

    let exeReader (read: RowReader) : LynxRow =

        let callMethodDynamically (instance: obj) (methodName: string) (args: obj[]) =
            let methodInfo = instance.GetType().GetMethod(methodName)
            methodInfo.Invoke(instance, args)

        let lynxRowType = typeof<LynxRow>
        let constructor = lynxRowType.GetConstructors().[0]
        let constructorArgs =
            lynxCols
            |> Array.map (fun {ColumnName = n; ColumnType = t} ->
                n
                |> deleteUpToFirstUnderscore
                |> fun columnName -> [| box columnName |]
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
        { grantYearStart = System.DateOnly(grantYear, 10, 1);
          lynxQuery = q
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