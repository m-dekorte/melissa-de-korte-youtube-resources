# Mapping Table: Column Renaming and Data Type Assignment in Power Query

This pattern lets you manage column renaming and data type assignment from a worksheet table, so people who do not use Power Query can update column names and types without opening the Power Query Editor.

- **Video:** [Watch on YouTube](https://youtu.be/b-1H0xYH2ZM)
- **LinkedIn:** [Connect with Melissa](https://www.linkedin.com/in/melissa-de-korte-64585884/)

---

## What this does

The query reads a mapping table, such as a table maintained in an Excel worksheet, and uses it to:

1. Rename source columns to friendly display names.
2. Apply the correct data type to each renamed column.

The person maintaining the mapping only needs to update the Excel table and refresh the query.

The mapping table has three columns:

| Column | Purpose |
|---|---|
| `oldColName` | The column name as it appears in the source data |
| `newColName` | The display name to use after renaming |
| `colType` | The data type to apply, written as a plain text string |

A validation drop-down on the `colType` column is recommended. It keeps the inputs consistent and helps prevent typing errors.

---

## About the demo file

In the demo, both `wsInputs` and `Source` are built inline in the same query using `Table.FromRows`. This keeps the example self-contained, so you can follow the logic without setting up external files or connections.

In practice, these two tables will often come from different places:

- **`wsInputs`** — a named Excel table in the current workbook, loaded as its own Power Query query
- **`Source`** — an external data source, such as another workbook, a CSV, a database, a SharePoint list, or similar

That separation can trigger a formula firewall error. See [Adapting for your own workbook](#adapting-for-your-own-workbook) for ways to handle it.

---

## The mapping table (`wsInputs`)

The worksheet table has exactly three columns.

**Example:**

| oldColName | newColName | colType |
|---|---|---|
| LastModified | Last Modified Date | date |
| Owner | Owner | text |
| NextReview | Review Date | date |
| WorkingTitle | Working Title | text |
| Revision | Revision Number | integer |
| Budget | Budget Amount | currency |
| Cost TD | Cost to date Amount | number |

---

## Supported type strings

Values in `colType` must exactly match one of the type strings below. Power Query is case-sensitive.

| Type string | M type applied |
|---|---|
| `any` | `Any.Type` |
| `date` | `Date.Type` |
| `time` | `Time.Type` |
| `datetime` | `DateTime.Type` |
| `logical` | `Logical.Type` |
| `number` | `Number.Type` |
| `currency` | `Currency.Type` |
| `integer` | `Int64.Type` |
| `percentage` | `Percentage.Type` |
| `text` | `Text.Type` |

An unrecognised string falls back to `Any.Type`. The column is left without a specific type, and no error is raised.

---

## The full M query

```m
let
    colTypes = [
        any = Any.Type,
        date = Date.Type,
        time = Time.Type,
        datetime = DateTime.Type,
        logical = Logical.Type,
        number = Number.Type,
        currency = Currency.Type,
        integer = Int64.Type,
        percentage = Percentage.Type,
        text = Text.Type
    ],
    wsInputs = Table.FromRows(
        {
            {"LastModified", "Last Modified Date", "date"},
            {"Owner", "Owner", "text"},
            {"NextReview", "Review Date", "date"},
            {"WorkingTitle", "Working Title", "text"},
            {"Revision", "Revision Number", "integer"},
            {"Budget", "Budget Amount", "currency"},
            {"Cost TD", "Cost to date Amount", "number"}
        },
        type table [oldColName = text, newColName = text, colType = text]
    ),
    Source = Table.FromRows(
        {
            {#date(2024, 1, 21), "Melissa", #date(2025, 1, 21), "To be determined", 1, 1000, 156.89},
            {#date(2011, 5, 2), "Sam", null, "Close encounters", 12, 500, 873.03}
        },
        Table.Column(wsInputs, "oldColName")
    ),
    colNames   = Table.ColumnNames(Source),
    ValidNames = Table.SelectRows(wsInputs, each List.Contains(colNames, [oldColName])),
    Renamed    = Table.RenameColumns(Source, List.Zip({ValidNames[oldColName], ValidNames[newColName]})),
    Transform  = Table.TransformColumnTypes(
        Renamed,
        List.Zip({
            ValidNames[newColName],
            List.Transform(ValidNames[colType], each Record.FieldOrDefault(colTypes, _, Any.Type))
        })
    )
in
    Transform
```

---

## Step-by-step walkthrough

### Step 1 — `colTypes`: build the type lookup

```m
colTypes = [
    any = Any.Type,
    date = Date.Type,
    time = Time.Type,
    datetime = DateTime.Type,
    logical = Logical.Type,
    number = Number.Type,
    currency = Currency.Type,
    integer = Int64.Type,
    percentage = Percentage.Type,
    text = Text.Type
],
```

**What it does:** Creates a record that maps plain text strings to type values.

`Table.TransformColumnTypes` does not accept text like `"date"`. It needs a type such as `Date.Type`. The `colTypes` record acts as the bridge between the worksheet input and the type values that M expects.

A record in M is a set of named fields. Each field has a name and a value. Here, the field names are the strings users enter in `wsInputs`, and the field values are the corresponding M types.

The field names are generalised identifiers, so they use more relaxed naming rules than regular M identifiers. Keywords such as `date`, `time`, and `text` are valid field names here without hash/pound-quoted notation, so `#"..."` is not needed.

---

### Step 2 — `wsInputs`: load the mapping table

```m
    wsInputs = Table.FromRows(
        {
            {"LastModified", "Last Modified Date", "date"},
            {"Owner", "Owner", "text"},
            {"NextReview", "Review Date", "date"},
            {"WorkingTitle", "Working Title", "text"},
            {"Revision", "Revision Number", "integer"},
            {"Budget", "Budget Amount", "currency"},
            {"Cost TD", "Cost to date Amount", "number"}
        },
        type table [oldColName = text, newColName = text, colType = text]
    ),
```

**What it does:** Builds the mapping table from a list of rows.

`Table.FromRows` takes a list of rows and a type declaration. Each inner list is one row. The type declaration sets the column names and specifies that all three columns contain text values.

In this demo, the mapping data is embedded directly in the query. In a real workbook, this step is usually replaced by a reference to an Excel table loaded as a separate query. In that setup, `wsInputs` comes from the current workbook, while `Source` often comes from an external file or system. That separation is what can trigger a formula firewall error. See [Adapting for your own workbook](#adapting-for-your-own-workbook).

---

### Step 3 — `Source`: load the raw data

```m
Source = Table.FromRows(
    {
        {#date(2024, 1, 21), "Melissa", #date(2025, 1, 21), "To be determined", 1, 1000, 156.89},
        {#date(2011, 5, 2), "Sam", null, "Close encounters", 12, 500, 873.03}
    },
    Table.Column(wsInputs, "oldColName")
),
```

**What it does:** Loads the source data using the original column names from `wsInputs`.

`Table.Column(wsInputs, "oldColName")` returns a list of all values in the `oldColName` column. That list becomes the column headers for the source table, so the data starts with the original source column names before the rename step runs.

In a real workbook, the `Table.FromRows(...)` call is replaced by the actual data connection to an external workbook, CSV, database, or other source.

---

### Step 4 — `colNames`: capture the current column names

```m
colNames = Table.ColumnNames(Source),
```

**What it does:** Returns the column names currently in `Source` as a list.

This list is used as a safeguard in the next step. Before the mapping is applied, the query checks which columns actually exist in the source data, instead of assuming every row in the mapping table is valid.

---

### Step 5 — `ValidNames`: filter the mapping to safe rows

```m
ValidNames = Table.SelectRows(wsInputs, each List.Contains(colNames, [oldColName])),
```

**What it does:** Keeps only the rows in `wsInputs` where `oldColName` matches a column that actually exists in `Source`.

`Table.SelectRows` filters the table row by row. The condition uses `List.Contains` to check whether the current row's `oldColName` appears in the `colNames` list.

Rows that do not match are removed before the rename and type steps run.

**Why this matters:** The mapping table is maintained by a person, and typos happen. A source column might also be renamed or removed. Without this check, a mismatch between `wsInputs` and the actual source data could cause `Table.RenameColumns` to return an error.

**What it does not do:** It does not notify the user when a row is skipped. The query continues, but the skipped column is not renamed or typed. For a more visible check, see [Adding a validation check](#adding-a-validation-check).

---

### Step 6 — `Renamed`: rename the columns

```m
Renamed = Table.RenameColumns(Source,
    List.Zip({ValidNames[oldColName], ValidNames[newColName]})
),
```

**What it does:** Renames each matched column using the old-name/new-name pairs from `ValidNames`.

`Table.RenameColumns` expects a list of rename operations. Each operation is a two-item list: `{oldName, newName}`.

`ValidNames[oldColName]` returns a list of original names. `ValidNames[newColName]` returns a list of display names. `List.Zip` pairs them by position, producing exactly the structure `Table.RenameColumns` needs.

Columns in `Source` that have no matching row in `ValidNames` are left unchanged.

---

### Step 7 — `Transform`: apply the data types

```m
Transform = Table.TransformColumnTypes(
    Renamed,
    List.Zip({
        ValidNames[newColName],
        List.Transform(ValidNames[colType], each Record.FieldOrDefault(colTypes, _, Any.Type))
    })
)
```

**What it does:** Applies the correct M type to each renamed column.

`Table.TransformColumnTypes` expects a list of `{columnName, type}` pairs. This step uses the same zipped-list pattern as the rename step.

The column names come from `ValidNames[newColName]`. This matters because the input table is `Renamed`, so the type step must refer to the new column names, not the original names.

The types come from `ValidNames[colType]`, which is a list of text strings. `List.Transform` goes through that list one item at a time. For each item, `Record.FieldOrDefault` looks up the text value as a field name in the `colTypes` record.

If the field exists, the matching M type is returned. If it does not exist, usually because of a typo or unsupported string, `Any.Type` is returned as the fallback. The query continues without error, but the affected column is left without a specific type.

---

## Adapting for your own workbook

The demo query is self-contained. Both `wsInputs` and `Source` are built from inline data so you can follow the logic without additional setup.

When you connect real data, the two tables may come from different places:

- **`wsInputs`** — loaded from a named Excel table in the current workbook as its own separate query
- **`Source`** — connected to an external file or data system

This is a common setup for this pattern. It is also the setup that can trigger the formula firewall.

---

### The formula firewall

When one Power Query query references another, and each connects to a different data source, Power Query may raise this error:

> *Formula.Firewall: Query references other queries or steps, so it may not directly access the data source. Please rebuild this combination.*

This is not a bug. It is a safeguard.

Power Query controls how data from different sources is combined during evaluation. Without that control, sensitive data from a private source could potentially be sent to a public one. The firewall flags that kind of combination and asks for direction.

There are two common ways to resolve it.

---

#### Option 1 — Embed the `wsInputs` code directly

Copy the full let-expression from the `wsInputs` query and assign it to a step inside the query that needs it. This is the same structure used in the demo file.

Because there is no cross-query reference, everything happens in the same partition, so the firewall has nothing to flag.

**Trade-off:** Each query that uses the mapping table gets its own independent copy. If the mapping changes, you must update it in every query separately.

**When to use it:** Use this when only one query relies on `wsInputs`, or when the mapping is stable and unlikely to change.

---

#### Option 2 — Convert `wsInputs` into a zero-parameter function

Rewrite the `wsInputs` query as a custom function that takes no parameters and returns the mapping table. To do that, add `() =>` above the `let` clause on line 1 of the `wsInputs` query:

```m
() =>
let
    wsInputs = Table.FromRows(
        {
            {"LastModified", "Last Modified Date", "date"},
            {"Owner", "Owner", "text"},
            {"NextReview", "Review Date", "date"},
            {"WorkingTitle", "Working Title", "text"},
            {"Revision", "Revision Number", "integer"},
            {"Budget", "Budget Amount", "currency"},
            {"Cost TD", "Cost to date Amount", "number"}
        },
        type table [oldColName = text, newColName = text, colType = text]
    )
in
    wsInputs  
```

Invoke it like this in any query that needs the mapping:

```m
wsInputs = wsInputs(),
```

When the mapping changes, you update the function once. Every query that calls it gets the updated version automatically.

**Trade-off:** This requires familiarity with custom functions in M. It is more maintainable, but it may be less familiar to people who are new to Power Query.

**When to use it:** Use this when multiple queries rely on the same mapping table and a single source of truth matters.

---

### Adding a validation check

The `ValidNames` step drops any row in `wsInputs` where `oldColName` does not match a column in `Source`. The query will not fail, but you may not notice that something was skipped.

To surface those skipped rows, duplicate the query, call it `DroppedRows`, and invert the filter condition for that step:

```m
ValidNames = Table.SelectRows(wsInputs, each not List.Contains(colNames, [oldColName]))
```

Return `DroppedRows` after the in-clause and load it as a separate query on a helper sheet. Any row that appears there was in the mapping table but did not match a column in `Source`, which usually means there is a typo or the source column no longer exists.

---

## Key M functions used

| Function | What it does in this query |
|---|---|
| `Table.FromRows` | Builds a table from a list of row values |
| `Table.Column` | Returns all values from one column as a list |
| `Table.ColumnNames` | Returns the column names of a table as a list |
| `Table.SelectRows` | Filters a table to rows where a condition is true |
| `List.Contains` | Checks whether a list contains a given value |
| `Table.RenameColumns` | Renames columns using a list of old/new name pairs |
| `Table.TransformColumnTypes` | Sets column types using a list of column/type pairs |
| `List.Zip` | Pairs two lists together by position |
| `List.Transform` | Applies a transformation to each item in a list |
| `Record.FieldOrDefault` | Looks up a field in a record; returns a default if not found |

---

## What to watch for

**Typos in `oldColName`** — the `ValidNames` step silently drops mismatched rows. If a mapping is not being applied, check the `oldColName` value against the actual column name in the source. Consider adding a `DroppedRows` step to surface these.

**Typos in `colType`** — an unrecognised type string falls back to `Any.Type`. The column is left without a specific type. Check that values are lowercase and match the supported list exactly.

**Case sensitivity** — column name matching is case-sensitive. `"Budget"` and `"budget"` are treated as different values. `oldColName` entries must match the source column names exactly.

**Using new names in `Transform`** — the `Transform` step works on `Renamed`, not `Source`. The column names passed to `Table.TransformColumnTypes` must be the new names, not the originals. The query handles this correctly, but keep this in mind if you adjust the step manually.

---

## Using the practice file

Copy the full M code script into a new blank query, replacing everything there with the code found here [The full M query](#the-full-m-query). Each step name corresponds to a section in this guide.

Try these to build understanding:

- Change a `colType` value to something unsupported — such as `xyz` — and check what type the column receives.
- Add a row to `wsInputs` with an `oldColName` that doesn't exist in `Source`. Watch whether the rename step errors or silently skips it.
- Deliberately misspell an `oldColName` entry and add a `DroppedRows` step to catch it.
- Add a column to `Source` that has no matching row in `wsInputs`. Observe what happens to that column through the query.

Review and test with your own data before using in a production process. Column names, data types, and source shapes vary.

## Resources

- **Full Step-by-Step Guide:** See sections above
- **Channel:** [YouTube](https://www.youtube.com/@melissa_de_korte)
- **Questions?** [Connect on LinkedIn](https://www.linkedin.com/in/melissa-de-korte-64585884/)
