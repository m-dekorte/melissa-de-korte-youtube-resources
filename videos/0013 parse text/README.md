# Parsing JSON Text Strings in Power Query

A step-by-step guide to turning JSON text strings into real lists, records and tables without relying on tricks.

- **Video:** [Watch on YouTube](https://youtu.be/91Az7O3O0Ws)
- **LinkedIn:** [Connect with Melissa](https://www.linkedin.com/in/melissa-de-korte-64585884/)

---

## The Problem

Some columns look like they contain lists or records. You can see the brackets, quotation marks and commas, but Power Query treats the cell as plain text, so there is nothing to expand.

A common workaround is to replace characters until it resembles M syntax...

---

## The Solution

That text is JSON, a standard format for structured data that often comes from APIs and exports. Power Query knows how to read it: `Json.Document` parses a JSON string and returns the value it describes.

Once the text is parsed, the shape of the result tells you which function turns it into a table:

| JSON text starts with | Parsed into | Typical next step |
|---|---|---|
| `[` (array) | List | `Table.FromColumns`, `Table.FromRows` or `Table.FromRecords` |
| `{` (object) | Record | Single record `Record.ToTable` or a list of records `Table.FromRecords` |

This guide builds three small reusable functions, one for each common shape, so you can apply the right one to your own data.

---

## Setup

### 1. Create the Parsed Query

In the Power Query Editor, create a new blank query named `Parsed`. It holds six JSON strings of increasing complexity and parses all of them in one step:

```powerquery
let
    Source = Table.FromValue( Record.FieldValues([
        Array = "[1,""A"",true,null,{""X"":10},[2,3]]",
        Object = "{""Name"":""Contoso"",""Count"":42,""Active"":true,""Note"":null}",
        EscapedQuotations = "{""Message"":""He said \""Hello\""."",""Description"":""A \""quoted\"" value""}",
        ArrayOfArrays = "[[1,2],[3,4],[5,6]]",
        ArrayOfObjects = "[{""ID"":1,""Name"":""A""},{""ID"":2,""Name"":""B""}]",
        ObjectOfArrays = "{""Tags"":[""Urgent"",""Online""],""LineIDs"":[501,502,503],""Flags"":[true,false]}"
    ])),
    JSON = Table.TransformColumns(Source,{},Json.Document),
    Result1 = JSON{0}[Value],
    Result2 = JSON{1}[Value],
    Result3 = JSON{2}[Value],
    Result4 = JSON{3}[Value],
    Result5 = JSON{4}[Value],
    Result6 = JSON{5}[Value]
in
    Result6
```

The query returns `Result6`. To inspect another parsed value, use the **Applied Steps** to select any of the intermediate values.

The entire column was parsed without writing code: select the column, then choose **Transform > Parse > JSON**.

### 2. Create fxArrayOfColArrays

Create a new blank query named `fxArrayOfColArrays`. Use this when every cell contains an array that represents **one column** of values that you want to convert into a table:

```powerquery
let
    /*
        Created by: Melissa de Korte
        Subscribe: https://www.youtube.com/@melissa_de_korte
        Resource: https://github.com/m-dekorte/melissa-de-korte-youtube-resources/blob/main/videos/0013%20parse%20text/README.md
    */
    fxArrayOfColArrays = Function.From(
        type function (fieldValue as text) as table, 
        each Table.FromColumns(
            List.Transform(_, Json.Document)
        )
    )
in
    fxArrayOfColArrays
```

### 3. Create fxArrayOfRowArrays

Create a new blank query named `fxArrayOfRowArrays`. Use this when every cell contains an array that represents **one row** of values that you want to convert into a table:

```powerquery
let
    /*
        Created by: Melissa de Korte
        Subscribe: https://www.youtube.com/@melissa_de_korte
        Resource: https://github.com/m-dekorte/melissa-de-korte-youtube-resources/blob/main/videos/0013%20parse%20text/README.md
    */
    fxArrayOfRowArrays = Function.From(
        type function (fieldValue as text) as table, 
        each Table.FromRows(
            List.Transform(_, Json.Document)
        )
    )
in
    fxArrayOfRowArrays
```

### 4. Create fxArrayOfObjects

Create a new blank query named `fxArrayOfObjects`. Use this when every cell contains an array of objects, where each object becomes a **row** in a new table:

```powerquery
let
    /*
        Created by: Melissa de Korte
        Subscribe: https://www.youtube.com/@melissa_de_korte
        Resource: https://github.com/m-dekorte/melissa-de-korte-youtube-resources/blob/main/videos/0013%20parse%20text/README.md
    */
    fxArrayOfObjects = Function.From(
        type function (fieldValue as text) as table, 
        each Table.FromRecords(List.Combine(
            List.Transform(_, Json.Document)
        ))
    )
in
    fxArrayOfObjects
```

### 5. Create the Sample Data Queries

Create two more blank queries. `SampleArrays` holds one row with two cells, each containing an array of text values. The first array has table names and the second has storage modes. Their positions line up:

```powerquery
let
    Source = Table.FromRows(
        {{
            "[""Fact Orders"",""Dim Product"",""Exchange Rates"",""Dim Date"",""Measures"",""Fact Returns"",""Dim Geography"",""Sales Targets"",""Dim Customer"",""Fact Sales""]", 
            "[""DirectQuery"",""Import"",""Import"",""Import"",""Import"",""Import"",""Import"",""Import"",""Import"",""DirectQuery""]"
        }}, 
        type table[tableName=text, tableStorageMode=text]
    )
in
    Source
```

`SampleObjects` holds one row where each cell contains an array of objects. The column names are reused from the previous sample for convenience:

```powerquery
let
    Source = Table.FromColumns(
        {
            {"[{""ID"":1,""Name"":""A""},{""ID"":2,""Name"":""B""}]"}, 
            {"[{""ID"":3,""Name"":""C""},{""ID"":4,""Name"":""D""}]"}
        }, 
        {"tableName", "tableStorageMode"}
    )
in
    Source
```

### 6. Invoke the Functions

Create three blank queries that call a sample and function. Each one adds a `tbl` column containing the resulting table.

Query `ArrayOfColArrays`:

```powerquery
let
    Source = SampleArrays,
    ArrayOfColArrays = Table.AddColumn(Source, "tbl", each fxArrayOfColArrays([tableName], [tableStorageMode]))
in
    ArrayOfColArrays
```

Query `ArrayOfRowArrays`:

```powerquery
let
    Source = SampleArrays,
    ArrayOfRowArrays = Table.AddColumn(Source, "tbl", each fxArrayOfRowArrays([tableName], [tableStorageMode]))
in
    ArrayOfRowArrays
```

Query `ArrayOfObjects`:

```powerquery
let
    Source = SampleObjects,
    ArrayOfObjects = Table.AddColumn(Source, "tbl", each fxArrayOfObjects([tableName], [tableStorageMode]))
in
    ArrayOfObjects
```

Click on the whitespace beside a `Table` value in the `tbl` column to preview a result, or use the expand button to bring its columns into the main table.

---

## Breakdown

### Json.Document

```powerquery
Json.Document(jsonText as any, optional encoding as nullable number) as any
```

| Argument | Meaning |
|---|---|
| `jsonText` | JSON as text, or as a binary value, so the same function can read a whole `.json` file |
| `encoding` | Optional. If omitted, UTF-8 is assumed |

The return type is `any` because the result depends on the JSON: an array gives a list, an object gives a record, and a scalar gives a single value.

### How JSON maps to M

| JSON | M |
|---|---|
| Array `[ ... ]` | List `{ ... }` |
| Object `{ ... }` | Record `[ ... ]` |
| String | Text |
| Number | Number |
| `true` / `false` | Logical |
| `null` | `null` |

Notice the brackets: a JSON array uses square brackets, but an M list uses curly brackets, and the reverse for objects and records.

Nested structures are handled automatically. In the `Parsed` query `Result1`, the array contains an object and another array, which become a nested record and a nested list.

Backslashes in front of quotation marks, as in `EscapedQuotations`, are JSON escape characters. They mark quotes that belong to the text, and the parser removes the backslash and keeps the quote.

---

### The Parsed query

```powerquery
Source = Table.FromValue(Record.FieldValues([ ... ]))
```

The inner record holds six named JSON strings. `Record.FieldValues` returns just the six values as a list, and `Table.FromValue` turns that list into a table with one row per item in a column named `Value`.

---

```powerquery
JSON = Table.TransformColumns(Source,{},Json.Document)
```

The second argument, the list of per-column operations, is empty. The third argument is the default transformation, which is applied to every column that has no specific operation. Here that means `Json.Document` runs on every cell of `Value`.

---

```powerquery
Result1 = JSON{0}[Value]
```

`JSON{0}` selects the first row, and `[Value]` accesses it's value. The result steps do the same for all rows:

| Step | JSON string | Parsed result |
|---|---|---|
| `Result1` | `Array` | List with six items, including a record and a list |
| `Result2` | `Object` | Record with four fields |
| `Result3` | `EscapedQuotations` | Record whose text values contain quotation marks |
| `Result4` | `ArrayOfArrays` | List of three lists |
| `Result5` | `ArrayOfObjects` | List of two records |
| `Result6` | `ObjectOfArrays` | Record whose fields each hold a list |

---

### Function.From

```powerquery
Function.From(
    type function (fieldValue as text) as table, 
    each ...
)
```

`Function.From` creates a function from a function type and a body. The body receives **all arguments as one list**, which is why `_`  will process every JSON strings that is passed to these functions.

It is what enables them to work with any number of arguments. You can pass one JSON text or ten, and the body simply transforms whatever list it receives.

---

### Parsing every argument

```powerquery
List.Transform(_, Json.Document)
```

Every JSON string in the argument list is parsed. If two strings are passed, the result is a list of two lists. The three functions differ only in how they turn that parsed structure into a table.

---

### Which table function to use

| Function | Treats each list or record as | Use when a cell holds |
|---|---|---|
| `Table.FromColumns` | a column | the values for one column |
| `Table.FromRows` | a row | the values for one row |
| `Table.FromRecords` | a row, with field names as columns | an array of objects |

Using `SampleArrays` shows why orientation matters:

| Function | Result |
|---|---|
| `fxArrayOfColArrays` | 10 rows and 2 columns, where each row pairs a name with its mode |
| `fxArrayOfRowArrays` | 2 rows and 10 columns, where row 1 holds the names and row 2 holds the modes |

For this data, the column version is the right one. The row version produces valid output, but in the wrong shape.

---

### List.Combine in fxArrayOfObjects

```powerquery
Table.FromRecords(List.Combine(List.Transform(_, Json.Document)))
```

Parsing an array of objects gives a list of records, so parsing several cells gives a list of lists of records. `Table.FromRecords` needs a single flat list, and `List.Combine` joins the inner lists into one.

---

### Calling the functions

```powerquery
Table.AddColumn(Source, "tbl", each fxArrayOfColArrays([tableName], [tableStorageMode]))
```

For each row, the function receives the values of the two columns and returns a table. `Table.AddColumn` stores that table in the new `tbl` column.

---

## Detailed example

This walkthrough follows `fxArrayOfObjects` using the `SampleObjects` query.

**Input:** one row with two text cells.

| tableName | tableStorageMode |
|---|---|
| `[{"ID":1,"Name":"A"},{"ID":2,"Name":"B"}]` | `[{"ID":3,"Name":"C"},{"ID":4,"Name":"D"}]` |

**Step 1.** `Table.AddColumn` call and pass it both fields `fxArrayOfObjects([tableName], [tableStorageMode])`.

**Step 2.** In the function body `List.Transform(_, Json.Document)` parses each text value. The result is a list of two lists, each holding two records:

```text
{
    { [ID = 1, Name = "A"], [ID = 2, Name = "B"] },
    { [ID = 3, Name = "C"], [ID = 4, Name = "D"] }
}
```

**Step 3.** `List.Combine` flattens it into one list of four records:

```text
{
    [ID = 1, Name = "A"],
    [ID = 2, Name = "B"],
    [ID = 3, Name = "C"],
    [ID = 4, Name = "D"]
}
```

**Step 4.** `Table.FromRecords` converts each record into a row, where field names map to column names:

| ID | Name |
|---:|---|
| 1 | A |
| 2 | B |
| 3 | C |
| 4 | D |

This table is the value stored in the `tbl` column of the `ArrayOfObjects` query.

---

## Adapt this for your data

### Pick the right function

| Your cell contains | Example | Use |
|---|---|---|
| A list of values that belongs in **one column** | `["a","b","c"]` | `fxArrayOfColArrays` |
| A list of values that belongs in **one row** | `["a","b","c"]` | `fxArrayOfRowArrays` |
| A list of objects | `[{"ID":1},{"ID":2}]` | `fxArrayOfObjects` |
| A single object | `{"Name":"Contoso"}` | `Json.Document`, then read the fields |

The same text can be valid for both array functions. Decide based on what the values represent in your data.

### Practical adjustments

- **Parse in place.** If you only need lists or records in a column, choose **Transform > Parse > JSON**. The functions in this guide are for when you want to build a table.
- **Rename columns.** `Table.FromColumns` and `Table.FromRows` name their columns `Column1`, `Column2` and so on. You can rename when or after expanding.
- **Set data types.** JSON has no date type, values arrive as text. Apply data types after parsing.
- **Expand the result.** Use the expand button on the `tbl` column, or replace `Table.AddColumn` with a `Table.Combine` pattern detailed next.

---

## Going further

> **Not covered in the video!** The two examples below build on the same idea but use a different approach, because the parsed values need to stay linked to an identifier in another column.

Real data often arrives with an identifier column next to the JSON text. Here, each row is a dataset with an ID, and the JSON describes the tables inside it. The goal is one output row per table, with the dataset ID repeated on every row.

### Example A

Create a query named `RAW_DATA1`. The `tableName` and `tableStorageMode` arrays line up by position:

```powerquery
let
    Source = Table.FromRows(
            {
                {
                    "1a2b3c4d-5e6f-4789-a012-3456789abcde",
                    "[""Fact Orders"",""Dim Product"",""Exchange Rates"",""Dim Date"",""Measures"",""Fact Returns"",""Dim Geography"",""Sales Targets"",""Dim Customer"",""Fact Sales""]",
                    "[""DirectQuery"",""Import"",""Import"",""Import"",""Import"",""Import"",""Import"",""Import"",""Import"",""DirectQuery""]"
                },
                {
                    "2b3c4d5e-6f70-4890-b123-456789abcdef",
                    "[""Forecast"",""Fact Revenue"",""Dim Department"",""Measures"",""Dim Calendar"",""Fact Budget"",""Dim Region"",""KPI Definitions"",""Fact Expenses"",""Dim Employee""]",
                    "[""Import"",""DirectQuery"",""Import"",""Import"",""Import"",""Import"",""Import"",""Import"",""DirectQuery"",""Import""]"
                },
                {
                    "3c4d5e6f-7081-4901-c234-56789abcdef0",
                    "[""Dim Warehouse"",""Fact Stock Movements"",""Measures"",""Dim Supplier"",""Reorder Levels"",""Fact Purchases"",""Dim Date"",""Inventory Targets"",""Fact Inventory"",""Dim Material""]",
                    "[""Import"",""DirectQuery"",""Import"",""Import"",""Import"",""DirectQuery"",""Import"",""Import"",""DirectQuery"",""Import""]"
                },
                {
                    "4d5e6f70-8192-4a12-d345-6789abcdef01",
                    "[""Fact Accounts Receivable"",""Dim Cost Center"",""Budget"",""Dim Date"",""Fact General Ledger"",""Measures"",""Dim Company"",""Financial Plan"",""Dim Account"",""Fact Accounts Payable""]",
                    "[""DirectQuery"",""Import"",""Import"",""Import"",""DirectQuery"",""Import"",""Import"",""Import"",""Import"",""DirectQuery""]"
                },
                {
                    "5e6f7081-92a3-4b23-e456-789abcdef012",
                    "[""Dim Channel"",""Fact Conversions"",""Marketing Budget"",""Dim Date"",""Measures"",""Fact Web Traffic"",""Dim Customer Segment"",""Campaign Targets"",""Fact Leads"",""Dim Campaign""]",
                    "[""Import"",""DirectQuery"",""Import"",""Import"",""Import"",""DirectQuery"",""Import"",""Import"",""DirectQuery"",""Import""]"
                }
            },
            type table[datasetId=text, tableName=text, tableStorageMode=text]
        )
in
    Source
```

Then create a query named `FINISHED1`:

```powerquery
let
    Source = RAW_DATA1,
    Result = Table.Combine(
        List.Transform(Table.ToRows(Source), each let
            d = _{0}, n = Json.Document(_{1}), m = Json.Document(_{2}) in
            Table.FromColumns(
                {
                    List.Repeat({d}, List.Count(n)), 
                    n, 
                    m
                },
                type table[datasetId=text, tableName=text, tableStorageMode=text]
            )
        )
    )
in
    Result
```

How it works:

| Step | What it does |
|---|---|
| `Table.ToRows(Source)` | Turns each table row into a list: `{datasetId, tableNames, storageModes}` |
| `d = _{0}` | The dataset ID of the current row |
| `n`, `m` | The two parsed arrays, with names in `n` and storage modes in `m` |
| `List.Repeat({d}, List.Count(n))` | Repeats the ID once per table name, so the column is as long as the others |
| `Table.FromColumns(..., type table[...])` | Builds a small table for this dataset and applies column names and types |
| `Table.Combine` | Stacks the five small tables into one |

Result: 50 rows, 10 for each dataset. The first rows look like this:

| datasetId | tableName | tableStorageMode |
|---|---|---|
| 1a2b3c4d-5e6f-4789-a012-3456789abcde | Fact Orders | DirectQuery |
| 1a2b3c4d-5e6f-4789-a012-3456789abcde | Dim Product | Import |
| 1a2b3c4d-5e6f-4789-a012-3456789abcde | Exchange Rates | Import |
| 1a2b3c4d-5e6f-4789-a012-3456789abcde | Dim Date | Import |
| ... | ... | ... |

### Example B

Here, the table names are the **keys** of a JSON object and the storage modes are the **values**. Create a query named `RAW_DATA2`:

```powerquery
let
    Source = Table.FromRows(
        {
            {
                "1a2b3c4d-5e6f-4789-a012-3456789abcde",
                "{""Fact Orders"":""DirectQuery"",""Dim Product"":""Import"",""Exchange Rates"":""Import"",""Dim Date"":""Import"",""Measures"":""Import"",""Fact Returns"":""Import"",""Dim Geography"":""Import"",""Sales Targets"":""Import"",""Dim Customer"":""Import"",""Fact Sales"":""DirectQuery""}"
            },
            {
                "2b3c4d5e-6f70-4890-b123-456789abcdef",
                "{""Forecast"":""Import"",""Fact Revenue"":""DirectQuery"",""Dim Department"":""Import"",""Measures"":""Import"",""Dim Calendar"":""Import"",""Fact Budget"":""Import"",""Dim Region"":""Import"",""KPI Definitions"":""Import"",""Fact Expenses"":""DirectQuery"",""Dim Employee"":""Import""}"
            },
            {
                "3c4d5e6f-7081-4901-c234-56789abcdef0",
                "{""Dim Warehouse"":""Import"",""Fact Stock Movements"":""DirectQuery"",""Measures"":""Import"",""Dim Supplier"":""Import"",""Reorder Levels"":""Import"",""Fact Purchases"":""DirectQuery"",""Dim Date"":""Import"",""Inventory Targets"":""Import"",""Fact Inventory"":""DirectQuery"",""Dim Material"":""Import""}"
            },
            {
                "4d5e6f70-8192-4a12-d345-6789abcdef01",
                "{""Fact Accounts Receivable"":""DirectQuery"",""Dim Cost Center"":""Import"",""Budget"":""Import"",""Dim Date"":""Import"",""Fact General Ledger"":""DirectQuery"",""Measures"":""Import"",""Dim Company"":""Import"",""Financial Plan"":""Import"",""Dim Account"":""Import"",""Fact Accounts Payable"":""DirectQuery""}"
            },
            {
                "5e6f7081-92a3-4b23-e456-789abcdef012",
                "{""Dim Channel"":""Import"",""Fact Conversions"":""DirectQuery"",""Marketing Budget"":""Import"",""Dim Date"":""Import"",""Measures"":""Import"",""Fact Web Traffic"":""DirectQuery"",""Dim Customer Segment"":""Import"",""Campaign Targets"":""Import"",""Fact Leads"":""DirectQuery"",""Dim Campaign"":""Import""}"
            }
        },
        type table[datasetId = text, tableStorage = text]
    )
in
    Source
```

Then create a query named `FINISHED2`:

```powerquery
let
    Source = RAW_DATA2,
    Result = Table.Combine(
        List.Transform(Table.ToRows(Source), each let 
            d = _{0}, j = Json.Document(_{1}) in
            Table.FromColumns(
                {
                    List.Repeat({d}, Record.FieldCount(j)),
                    Record.FieldNames(j), 
                    Record.FieldValues(j)
                },
                type table[datasetId=text, tableName=text, tableStorageMode=text]
            )
        )
    )
in
    Result
```

How it differs from Example A:

| Step | What it does |
|---|---|
| `j = Json.Document(_{1})` | Parses the object into a record |
| `Record.FieldCount(j)` | Counts the fields, which sets how often the ID is repeated |
| `Record.FieldNames(j)` | Returns the keys as a list, used for `tableName` |
| `Record.FieldValues(j)` | Returns the values as a list, used for `tableStorageMode` |

The output has the same shape as Example A: 50 rows with `datasetId`, `tableName` and `tableStorageMode`.

---

## Notes

This solution assumes:

- The text in each cell is valid JSON. Invalid JSON, empty text or `null` makes `Json.Document` return an error, so clean or filter those rows first, or protect the expression using `try ... otherwise`.
- When you pass several JSON values to one function, the arrays have the same length. Otherwise the rows will not line up the way you expect.
- For `fxArrayOfObjects`, all objects share the same field names.
- JSON object keys are unique. A record cannot hold duplicate field names, the same applies to M.
- The `Function.From` function type declares one parameter, `fieldValue as text`, but these samples pass two. The function body receives a list with argument values, so it can handle any number of arguments that are passed. 

If your queries or columns have different names, update these parts:

```powerquery
Source = SampleArrays
```

```powerquery
fxArrayOfColArrays([tableName], [tableStorageMode])
```

If the JSON text sits in a column named something else, replace `[tableName]` and `[tableStorageMode]` with your own column references.

---

## Tags

`Power Query` · `M code` · `JSON` · `Json.Document` · `Function.From` · `Table.FromColumns` · `Table.FromRows` · `Table.FromRecords` · `List.Combine` · `Table.FromValue` · `Table.TransformColumns` · `List.Repeat` · `Table.Combine` · `Record.FieldNames` · `Excel` · `Power BI`
