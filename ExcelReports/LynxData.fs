module ExcelReports.LynxData

(*

#load "ExcelReports/OIBTypes.fs";;
open ExcelReports.OIBTypes;;

#r "nuget: Npgsql.FSharp, 5.7.0";;
#load "ExcelReports/LynxData.fs";;
open ExcelReports.LynxData;;

*)

open System
open System.Reflection
open Npgsql.FSharp
open System.Text.RegularExpressions

open OIBTypes

let (|GetType|_|) (ct: Type) (t: Type) =
    match Regex("Option").Match(t.FullName).Success with
    | true ->
        Regex("\[(System.*?),").Match(t.FullName).Groups.[1].Value
        |> fun (extractedType: string) -> (extractedType = ct.FullName)
        |> fun (isMatch: bool) -> ( isMatch, "option" )
        |> Some
    | false ->
        Some <| ( (t.FullName = ct.FullName), "" )

// let fieldNamesAndTypes = getRecordFieldNamesAndTypes<LynxRow>()

(* HISTORICAL REMINDERS

    1.

        (r : R).GetType() = typeof<R>

    2. Just copy into `dotnet fsi`

        type Ia = abstract member ia: unit -> unit
        type Ib = abstract member ib: unit -> Ia
        type X = { miez : string } interface Ia with member this.ia () = () end interface Ib with member this.ib () = this end

        let ret<'T> (this : 'T) = this

        type Y = { miez : string } interface Ia with member this.ia () = () end interface Ib with member this.ib () = (ret<Y> this) end

        (( { miez = "lofa" } : Y ) :> Ib).ib ();;

    3.

        type I = abstract member ia: unit -> string abstract member ib: unit -> string

        type X = Lofa interface I with member this.ia () = "miez" member this.ib () = (this :> I).ia () + " vmi"
        (Lofa :> I).ib ()

        let f (this : I) = I.ia () + " balabab"
        type Y = Miez interface I with member this.ia () = "lofa" member this.ib () = f this
        (Miez :> I).ib ();;

*)

// TODO 2023-12-01_1342
//      Remove the hard-coded password.
// let connectionString = "postgres://postgres:password@192.168.64.4:5432/lynx"

// Record types that can be used to build SQL queries.
type ISQLQueryColumnable =
    abstract member _ignore: unit -> unit

(*

*)
type LynxColumnSpecification = {
    ColumnName : string;
    ColumnType : System.Type;
}

type LynxQueryRowSpecification = LynxColumnSpecification array

(*
    Takes a record TYPE (not instance!) that represents a returned SQL query row (and one that implements the `ISQLQueryColumnable` interface) and return the `LynxColumnSpecification` representation for each column of the input type specification.

    Used for:

        1. Generating the columns in `SELECT ..columns.. FROM` part of the SQL query.

        2. Constructing the `RowReader` (nee `NpgsqlDataReader` in Npgsql) type input for Npgsql.FSharp's `Sql.execute` function to do a SQL query.

    NOTE
        Cannot be moved inside a function because of  FS0665.
        ("Explicit type parameters may only be used on module
         or member bindings")
*)
let getRecordFieldNamesAndTypes<'T when 'T :> ISQLQueryColumnable>
    : LynxQueryRowSpecification
     =

    let toLynxColumn (fieldInfo: FieldInfo) : LynxColumnSpecification =
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

    typeof<'T>.GetFields(BindingFlags.Public ||| BindingFlags.Instance)
    |> Array.map toLynxColumn

type SqlAliases = string list
type SqlAliasFor = ForQueryColumn | ForRowReader

let processSqlAliases
    (sqlAliases: SqlAliases)
    (sqlAliasFor: SqlAliasFor)
    (lynxColumnName: string)
    : string
    =

    let deleteUpToFirstUnderscore (str: string) =
        let index = str.IndexOf('_')
        if index >= 0 then
            str.Substring(index + 1)
        else
            str

    let isColumnAliased =
        List.contains lynxColumnName sqlAliases

    match (sqlAliasFor, isColumnAliased) with

    | (ForQueryColumn, true)  -> lynxColumnName + " AS " + lynxColumnName
    | (ForQueryColumn, false) -> lynxColumnName

    | (ForRowReader, true)  -> lynxColumnName
    | (ForRowReader, false) -> deleteUpToFirstUnderscore lynxColumnName

// Generic helper to implement the `ISQLQueryColumnable` interface.
let buildQueryColumns<'T when 'T :> ISQLQueryColumnable>
    (sqlAliases: SqlAliases)
    : string
    =

    let replaceFirstOccurrence (oldValue: char, newValue: char) (str: string) =
        let index = str.IndexOf(oldValue)
        if index >= 0 then
            str.Remove(index, 1).Insert(index, newValue.ToString())
        else
            str

    // SELECT columns generated from `ISQLQueryColumnable` types.
    getRecordFieldNamesAndTypes<'T>
    |> Array.map
        ( fun ({ColumnName = n; ColumnType = t}: LynxColumnSpecification) ->
            n
            |> processSqlAliases sqlAliases ForQueryColumn
            |> replaceFirstOccurrence ('_', '.')
        )
    |> String.concat ", "

// see TODO 2024-02-19_1348 rename_lynx_prefix_oib_report
type QuarterlyReportQueryRow =
// type LynxRow =
    {
        // NOTE 2023-12-14_1221 Why are most record fields options?
        // Because  the LYNX  database  tables  have almost  no
        // constraints (practically used as an Excel workbook),
        // so it is  just safer to treat any  returned value as
        // invalid.

        // NOTE Naming convention of the record fields
        //
        //      <table_alias>_<column_name>
        //
        //      where `<table_alias>`-es are defined
        //      in the  `quarterly7OBReportQuery` function in the
        //      `joins` variable.

        // NOTE How to use SQL aliases?
        //
        //      See NOTE "add_SQL_aliases" before `quarterly7OBReportQuery`
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
        plan_employment_outcomes     : string option;
        plan_living_plan_progress    : string option;
        plan_plan_date               : System.DateOnly option;
        plan_modified                : System.DateTime option;

        mostRecentAddress_id       :    int; // aliased! see NOTE "add_SQL_aliases"
        mostRecentAddress_modified : System.DateTime option;
        mostRecentAddress_county   : string option;
    }

    interface ISQLQueryColumnable with
        member this._ignore () = ()

(* The original SQL query:

    SELECT
        assignment.id, assignment.assignment_date,
        contact.last_name, contact.first_name,
        instructor.last_name, instructor.first_name,
        assignedby.last_name, assignedby.first_name,
        assignment.assignment_status
    FROM
            lynx_assignment AS assignment
        JOIN lynx_contact    AS contact    ON    contact.id = assignment.contact_id
        JOIN auth_user       AS instructor ON instructor.id = assignment.instructor_id
        JOIN auth_user       AS assignedby ON assignedby.id = assignment.user_id
    ORDER BY assignment.assignment_date DESC
*)
type AssignmentQueryRow =
    {
        assignment_id                : int;
        contact_last_name            : string option;
        contact_first_name           : string option;
        instructor_last_name         : string option;
        instructor_first_name        : string option;
        assignedby_last_name         : string option;
        assignedby_first_name        : string option;
        assignment_assignment_status : string option;
        assignment_assignment_date   : System.DateTime option;
    }

    interface ISQLQueryColumnable with
        member this._ignore () = ()

// type LynxColumns =
//     {
//         oib_report_query : LynxQueryRowSpecification;
//         assignment_query : LynxQueryRowSpecification;
//     }

// let lynxColumns =
//     {
//         oib_report_query = getRecordFieldNamesAndTypes<LynxRow, LynxColumnSpecification> toLynxColumn
//     ;   assignment_query = getRecordFieldNamesAndTypes<AssignmentQueryRow, LynxColumnSpecification> toLynxColumn
//     }

// see TODO 2024-02-19_1348 rename_lynx_prefix_oib_report
// type Quarterly7OBReportQuery = QuarterlyReportQueryRow list

// TODO there are two quarterly reports (7OB and non-7OB), so this type should accomodate both
type OIBQuarterlyReportData = {
    grantYearStart : System.DateOnly;
    grantYearEnd   : System.DateOnly;
    // quarterly7OBReportQuery : Quarterly7OBReportQuery;
    lynxData : ISQLQueryColumnable list;
}

type Quarter =
    | Q1
    | Q2
    | Q3
    | Q4

    interface System.IFormattable with
        member this.ToString(_format: string, _formatProvider: System.IFormatProvider) =
            match this with
            | Q1 -> "Q1"
            | Q2 -> "Q2"
            | Q3 -> "Q3"
            | Q4 -> "Q4"

type RowReaderBuilder =
    LynxQueryRowSpecification -> SqlAliases -> RowReader -> ISQLQueryColumnable

let callMethodDynamically (instance: obj) (methodName: string) (args: obj[]) =
    let methodInfo = instance.GetType().GetMethod(methodName)
    methodInfo.Invoke(instance, args)

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

let rowReaderBuilder<'T when 'T :> ISQLQueryColumnable>
    // (rowSpec: LynxQueryRowSpecification)
    (sqlAliases: SqlAliases)
    (read: RowReader)
    : ISQLQueryColumnable
    =
    // let getRecordFieldNamesAndTypes<'T when 'T :> ISQLQueryColumnable>

    // printfn "rowReaderBuilder"
    let rowSpec: LynxQueryRowSpecification =
        getRecordFieldNamesAndTypes<'T>

    let lynxRowType = typeof<'T>
    let constructor = lynxRowType.GetConstructors().[0]
    let constructorArgs: obj array =
        rowSpec
        |> Array.map (fun {ColumnName = n; ColumnType = t} ->
            // printfn $"COLUMN NAME: {n}"
            n
            |> processSqlAliases sqlAliases ForRowReader
            |> fun columnName -> [| box columnName |]
            |> callMethodDynamically read (typeToRowReaderMember t)
            |> box
        )

    constructor.Invoke(constructorArgs) :?> ISQLQueryColumnable

    // The output is something along these lines:
    //
    //      {
    //          ContactID = read.int "id";
    //          LastName = read.text "last_name";
    //          ...
    //      }


    // See more examples https://github.com/Zaid-Ajaj/Npgsql.FSharp?tab=readme-ov-file#sqlexecute-execute-query-and-read-results-as-table-then-map-the-results
    // but here's a snippet of the end result:
    //
    //   {
    //       ContactID = read.int "id";
    //       LastName = read.text "last_name";
    //       ...
    //   }

(*
    Function parameters:

    + `connectionString`

        The PostgreSQL connection URI as a string.
        https://www.postgresql.org/docs/14/libpq-connect.html#id-1.7.3.8.3.6

    + `sqlQueryStringAfterFROM`

        The  SQL  query string  after the `SELECT (..columns..) FROM` part. See
        `rowSpec` parameter below of how `..columns..` is generated.

        TODO: add example string

    + `rowSpec`

        The  `rowSpec`  parameter   will  always  be  provided   by  using  the
        `getRecordFieldNamesAndTypes`  function  as  it  can't  be placed inside
        the `queryBuilder` function because of  FS0665. (I got excited about the
        snippet  in  https://stackoverflow.com/a/32345373/1498178, but,  in  the
        end, it is  the same function as  `getRecordFieldNamesAndTypes`, and has
        to be called with an explicit type parameter.)

        Example:

            getRecordFieldNamesAndTypes<AssignmentQueryRow>
            |> queryBuilder connectionString sqlQueryStringAfterFROM
*)
type QueryBuilderParameters =
    { connectionString: string
    ; sqlQueryStringAfterFROM: string
    ; sqlAliases: SqlAliases
    // ; rowSpec: LynxQueryRowSpecification
 // Made this implicit as it is not supposed to change
 // ; rowReaderBuilder: RowReaderBuilder
    }

let queryBuilder<'T when 'T :> ISQLQueryColumnable>
    (args: QueryBuilderParameters)
    : ISQLQueryColumnable list
    =

    let queryColumns: string =
        buildQueryColumns<'T> args.sqlAliases

    let queryString =
        "SELECT " + queryColumns + " FROM " + args.sqlQueryStringAfterFROM

    printfn $"QUERY: {queryString}"

    args.connectionString
    |> Sql.connect
    |> Sql.query queryString
    |> Sql.execute ( rowReaderBuilder<'T> args.sqlAliases )

type QuarterlyOIBReportType =
    | OIB_7OB
    | OIB_Non7OB

let assignmentReportQuery
    (connectionString: string)
    (reportType: QuarterlyOIBReportType)
    : ISQLQueryColumnable list
    =

    let sqlQueryStringAfterFROM =
        $"""
             lynx_{if (reportType = OIB_7OB) then "" else "sip1854"}assignment AS assignment
        JOIN lynx_contact    AS contact    ON    contact.id = assignment.contact_id
        JOIN auth_user       AS instructor ON instructor.id = assignment.instructor_id
        JOIN auth_user       AS assignedby ON assignedby.id = assignment.user_id
        """

    queryBuilder<AssignmentQueryRow>
        { connectionString = connectionString
        ; sqlQueryStringAfterFROM = sqlQueryStringAfterFROM
        ; sqlAliases =
            [  "contact_last_name"
            ; "contact_first_name"
            ;  "instructor_last_name"
            ; "instructor_first_name"
            ;  "assignedby_last_name"
            ; "assignedby_first_name"
            ]
        }

let quarterToStartAndEndDates (q: Quarter) (grantYear: int) =
    let startDate =
        match q with
        | Q1 -> System.DateOnly (grantYear    , 10, 1)
        | Q2 -> System.DateOnly (grantYear + 1,  1, 1)
        | Q3 -> System.DateOnly (grantYear + 1,  4, 1)
        | Q4 -> System.DateOnly (grantYear + 1,  7, 1)

    ( startDate
    , startDate.AddMonths(3).AddDays(-1)
    )

// TODO There is also a non-7OB report, so this function should be generalized to support both (either by splitting out the generic parts or by adding an extra parameter; all the non-7OB parts in LYNX are the same as the 7OB ones but containing the "1854" label somewhere).
let quarterlyReportQuery
    (connectionString: string)
    (reportType: QuarterlyOIBReportType)
    (quarter: Quarter)
    (grantYear: int)
    : OIBQuarterlyReportData
    =

    let grantYearStart : DateOnly = quarterToStartAndEndDates Q1 grantYear |> fst
    let grantYearEnd   : DateOnly = quarterToStartAndEndDates Q4 grantYear |> snd

    // `quarterStart` will always be `grantYearStart`
    // because the OIB report is a cumulative one
    let quarterStart : DateOnly = grantYearStart
    let quarterEnd   : DateOnly = quarterToStartAndEndDates quarter grantYear |> snd

    let is7OB: bool = reportType = OIB_7OB
    let noteTable = if is7OB then "lynx_sipnote" else "lynx_sip1854note"
    let planTable = if is7OB then "lynx_sipplan" else "lynx_sip1854plan"

    let joins =
        $"""
             {noteTable}  AS note
        JOIN lynx_contact AS contact ON    note.contact_id = contact.id
        JOIN {planTable}  AS plan    ON   note.sip_plan_id = plan.id
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

    let whereClause =
        $"""
        WHERE note.note_date >= '{quarterStart.ToString()}'::date
          AND note.note_date <  '{quarterEnd.ToString()}'::date
        """

    // NOTE 2023-12-01_1347 Should be irrelevant.
    // let groupByClause = "GROUP BY " + queryColumns
    // let orderByClause = "ORDER BY CONCAT(c.last_name, ', ', c.first_name)"

    queryBuilder<QuarterlyReportQueryRow>
        { connectionString = connectionString
        ; sqlQueryStringAfterFROM = joins + whereClause
        ; sqlAliases =
            [ "contact_id"
            ; "intake_id"
            ; "note_id"
            ; "plan_id"
            ; "plan_modified"
            ]
        }
    |> fun (q: ISQLQueryColumnable list) ->
        { grantYearStart = grantYearStart
        // Only used in `getAgeAtApplication` in `Library.fs`
        ; grantYearEnd = grantYearEnd
        ; lynxData = q
        }