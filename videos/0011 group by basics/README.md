# Group By Basics: Percentage of Group Total

Calculating a percentage within each group in Power Query.

- Video: [Watch on YouTube](https://www.youtube.com/watch?v=3xwSoSiUGsE)
- LinkedIn: [Connect with Melissa](https://www.linkedin.com/in/melissa-de-korte/)
- Challenge creator: [Myles Arnott on LinkedIn](https://www.linkedin.com/in/mylesarnott/)
- Challenge post: [View LinkedIn post](https://www.linkedin.com/posts/mylesarnott_excel-powerquery-ugcPost-7501181039852314624-fHw3/)

---

## Scenario

Each row contains a sighting count for a species at a location. The goal is to calculate what percentage each species represents within its location, not across the entire dataset.

The core difficulty is that the denominator for the percentage changes per group. For example, Hales Wood has a total of 61 sightings, while Quantock Wood has 180. Each row needs to divide by its own group total.

This document walks through two solutions:

- **Solution A** uses Group By with All Rows to calculate a group total, expand the details, and divide.
- **Solution B** uses Group By with a nested `Table.AddColumn` to calculate the percentage.

## Challenge rules

Calculate the percentage of sightings for each species within its location:

1. Group the data by `Location`.
2. Calculate the total number of sightings per location.
3. Divide each row's `No. of Sightings` by its location total.
4. Return the result as a percentage.

---

## Source data

Sample table:

| Location | Species | No. of Sightings |
|---|---|---:|
| `Hales Wood` | `Grey` | `5` |
| `Hales Wood` | `Red` | `56` |
| `Green Wood` | `Grey` | `214` |
| `Green Wood` | `Red` | `20` |
| `Hake Wood` | `Grey` | `7` |
| `Hake Wood` | `Red` | `40` |
| `Forest of Dean` | `Red` | `54` |
| `Forest of Dean` | `Grey` | `210` |
| `Pile Wood` | `Grey` | `4` |
| `Pile Wood` | `Red` | `60` |
| `Quantock Wood` | `Grey` | `180` |
| `Quantock Wood` | `Red` | `0` |

Expected result:

| Location | Species | No. of Sightings | Location % |
|---|---|---:|---:|
| `Hales Wood` | `Grey` | `5` | `8.20%` |
| `Hales Wood` | `Red` | `56` | `91.80%` |
| `Green Wood` | `Grey` | `214` | `91.45%` |
| `Green Wood` | `Red` | `20` | `8.55%` |
| `Hake Wood` | `Grey` | `7` | `14.89%` |
| `Hake Wood` | `Red` | `40` | `85.11%` |
| `Forest of Dean` | `Red` | `54` | `20.45%` |
| `Forest of Dean` | `Grey` | `210` | `79.55%` |
| `Pile Wood` | `Grey` | `4` | `6.25%` |
| `Pile Wood` | `Red` | `60` | `93.75%` |
| `Quantock Wood` | `Grey` | `180` | `100.00%` |
| `Quantock Wood` | `Red` | `0` | `0.00%` |

## Raw data query

```powerquery
let
    Source = Table.FromRows(
        {
            {"Hales Wood", "Grey", 5},
            {"Hales Wood", "Red", 56},
            {"Green Wood", "Grey", 214},
            {"Green Wood", "Red", 20},
            {"Hake Wood", "Grey", 7},
            {"Hake Wood", "Red", 40},
            {"Forest of Dean", "Red", 54},
            {"Forest of Dean", "Grey", 210},
            {"Pile Wood", "Grey", 4},
            {"Pile Wood", "Red", 60},
            {"Quantock Wood", "Grey", 180},
            {"Quantock Wood", "Red", 0}
        },
        type table [Location=text, Species=text, No. of Sightings=Int64.Type]
    )
in
    Source
```

This creates the sample dataset. Name this query `RAW_DATA` so both solutions can reference it.

---

## Solution A: Group By basics, no M coding

This solution uses the Power Query interface to group the data, keep the detail rows, expand them, and calculate the percentage.

### Power Query M solution

```powerquery
let
    Source = RAW_DATA,
    GroupRows = Table.Group(
        Source,
        {"Location"},
        {
            {"Details", each _, type table [Location=text, Species=text, No. of Sightings=number]},
            {"Group Total", each List.Sum([No. of Sightings]), type number}
        }
    ),
    ExpandDetails = Table.ExpandTableColumn(GroupRows, "Details", {"Species", "No. of Sightings"}),
    InsertDivision = Table.AddColumn(ExpandDetails, "Location %", each [No. of Sightings] / [Group Total], Percentage.Type),
    RemoveCols = Table.RemoveColumns(InsertDivision, {"Group Total"})
in
    RemoveCols
```

### Breakdown

The solution groups by Location, creates two aggregations, expands the details, and calculates the percentage.

---

```powerquery
Source = RAW_DATA
```

This references the raw data query created earlier.

---

```powerquery
Table.Group(
    Source,
    {"Location"},
    {
        {"Details", each _, type table [...]},
        {"Group Total", each List.Sum([No. of Sightings]), type number}
    }
)
```

`Table.Group` groups the source table by `Location`.

Two aggregations are created:

| Aggregation | Operation | Result |
|---|---|---|
| `Details` | `each _` (All Rows) | a nested table containing all rows for that group |
| `Group Total` | `List.Sum([No. of Sightings])` | the sum of sightings for that group |

After this step, the table has one row per location:

| Location | Details | Group Total |
|---|---|---:|
| `Hales Wood` | `Table` | `61` |
| `Green Wood` | `Table` | `234` |
| `Hake Wood` | `Table` | `47` |
| `Forest of Dean` | `Table` | `264` |
| `Pile Wood` | `Table` | `64` |
| `Quantock Wood` | `Table` | `180` |

---

```powerquery
each _
```

This is the key to keeping the detail rows.

In the context of `Table.Group`, `each _` means "return all the rows that belong to this group as a nested table."

This is equivalent to selecting **All Rows** in the Advanced Group By dialog in the Power Query interface.

For example, the nested table for Hales Wood contains:

| Location | Species | No. of Sightings |
|---|---|---:|
| `Hales Wood` | `Grey` | `5` |
| `Hales Wood` | `Red` | `56` |

The Group Total for the same row is `61`.

This means the denominator and the detail rows now live on the same row.

---

```powerquery
each List.Sum([No. of Sightings])
```

This sums the `No. of Sightings` column within each group.

Inside a `Table.Group` aggregation, `[No. of Sightings]` refers to the column values from the grouped rows, not from the entire source table.

For Hales Wood, this returns `5 + 56 = 61`.

---

```powerquery
ExpandDetails = Table.ExpandTableColumn(GroupRows, "Details", {"Species", "No. of Sightings"})
```

`Table.ExpandTableColumn` brings the rows from the nested `Details` table back into the main table.

Only `Species` and `No. of Sightings` are expanded because `Location` already exists in the outer table.

After expanding, the table returns to one row per species per location, but with the `Group Total` column still present:

| Location | Group Total | Species | No. of Sightings |
|---|---:|---|---:|
| `Hales Wood` | `61` | `Grey` | `5` |
| `Hales Wood` | `61` | `Red` | `56` |

The Group Total value is repeated for every row that belongs to the same location.

---

```powerquery
InsertDivision = Table.AddColumn(
    ExpandDetails,
    "Location %",
    each [No. of Sightings] / [Group Total],
    Percentage.Type
)
```

To create a new column called `Location %`. Select the [No. of Sightings] column, hold down CTRL and select [Group Total]. The order matters because we need to divide [No. of Sightings] by [Group Total].

On the ribbon go to: Add Column/ From Number/ Standard/ Divide

Update the column name (Division > `Location %`) and the ascribed type (type number > `Percentage.Type`)

For example, for the Grey squirrel row in Hales Wood:

```text
5 / 61 = 0.0820 → 8.20%
```

The type is set to `Percentage.Type` so the result displays as a percentage.

---

```powerquery
RemoveCols = Table.RemoveColumns(InsertDivision, {"Group Total"})
```

The `Group Total` column was only needed for the division step.

Once the percentage is calculated, the helper column is removed from the final result.

---

## Solution B: Group By advanced with M coding

This solution calculates the percentage inside the grouping step.

### Power Query M solution

```powerquery
let
    Source = RAW_DATA,
    GroupBy = Table.Group(
        Source,
        {"Location"},
        {
            {"Sightings", each Table.AddColumn(
                _,
                "newFields",
                (n) => [
                    total = List.Sum([No. of Sightings]),
                    pct = Value.Divide(n[No. of Sightings], total, Precision.Decimal)
                ]
            ), type table [Location=text, Species=text, No. of Sightings=number, newFields=[total=number, pct=Percentage.Type]]}
        }
    ),
    ExpandSightings = Table.ExpandTableColumn(GroupBy, "Sightings", {"Species", "No. of Sightings", "newFields"}),
    ExpandFields = Table.ExpandRecordColumn(ExpandSightings, "newFields", {"pct"}, {"Location %"})
in
    ExpandFields
```

### Breakdown

The solution performs the calculation inside the grouping step by adding a column to the inner table. This requires understanding how scope works inside `Table.Group`.

---

```powerquery
Table.Group(
    Source,
    {"Location"},
    {
        {"Sightings", each Table.AddColumn(...), type table [...]}
    }
)
```

Instead of creating separate aggregations like Solution A, this solution creates a single aggregation called `Sightings`.

That aggregation does not just return the grouped rows. It returns a modified version of the grouped rows with new calculated fields added.

---

### Scope: the inner table and the inner row

This is the most important concept in this solution.

Inside a `Table.Group` aggregation, there are two layers of scope:

| Scope | What it refers to | How to access |
|---|---|---|
| **Inner table** (the group) | all rows belonging to the current group | `_` inside `each`, or `[ColumnName]` |
| **Inner row** | one row inside the inner table | named parameter, such as `(n)` |

When `Table.AddColumn` is called inside the aggregation, it operates on the table. But `Table.AddColumn` also introduces its own context for each row in that table.

```powerquery
each Table.AddColumn(
    _,                                  // inner table (all rows in this group)
    "newFields",
    (n) => [                            // n = current row in the table
        total = List.Sum([No. of Sightings]),   // [No. of Sightings] = table column
        pct = Value.Divide(n[No. of Sightings], total, Precision.Decimal)  // n[No. of Sightings] = current row
    ]
)
```

---

```powerquery
each Table.AddColumn(_, "newFields", (n) => ...)
```

The `each ... _` receives the table for the current group.

`Table.AddColumn` then loops through each row in that table and adds a new column called `newFields`.

The row-level function (columnGenerator) uses a parameter named `n` instead of `_` so that the row and the group table can be accessed independently.

---

```powerquery
total = List.Sum([No. of Sightings])
```

Here, `[No. of Sightings]` refers to the column from the **inner table**, meaning all sighting values for the current group.

For Hales Wood, this evaluates to:

```text
List.Sum({5, 56}) = 61
```

This works because the record expression `[ total = ..., pct = ... ]` is defined inside a columnGenerator function `(n) =>` called by `Table.AddColumn`. Inside this record, `[No. of Sightings]` without a prefix resolves to the column of the inner table from the enclosing `each _` scope, not to a field on the current row.

---

```powerquery
pct = Value.Divide(n[No. of Sightings], total, Precision.Decimal)
```

Here, `n[No. of Sightings]` refers to the sighting count from the **current row**.

The named parameter `n` gives access to the individual row, so `n[No. of Sightings]` returns the value for that specific species.

For the Grey squirrel row in Hales Wood:

```text
Value.Divide(5, 61, Precision.Decimal) = 0.0820 → 8.20%
```

`Value.Divide` is used here instead of the `/` operator. The third argument, `Precision.Decimal`, avoids floating-point rounding issues by using decimal precision.

---

### Why the named parameter matters

If the columnGenerator function used `each` instead of `(n) =>`, only the closest scope from its perspective would survive. From the columnGenerator's perspective that's the individual row - access to the grouped table would be lost.

With `each`, the underscore `_` refers to the current row, and `[No. of Sightings]` would also refer to the current row's value, not the table's column.

Using a unique parameter name like `n` keeps the two scopes separate:

| Expression | Resolves to |
|---|---|
| `[No. of Sightings]` | the column from the inner table (group scope) |
| `n[No. of Sightings]` | the field from the current row |

This is what makes the percentage calculation possible inside the grouping step.

---

```powerquery
(n) => [
    total = List.Sum([No. of Sightings]),
    pct = Value.Divide(n[No. of Sightings], total, Precision.Decimal)
]
```

The function returns a record with two fields:

| Field | Value |
|---|---|
| `total` | the group total |
| `pct` | the percentage for the current row |

This record is stored in the `newFields` column.

---

```powerquery
ExpandSightings = Table.ExpandTableColumn(
    GroupBy,
    "Sightings",
    {"Species", "No. of Sightings", "newFields"}
)
```

This expands the nested `Sightings` table back into the main table.

After this step, each row has the original columns plus a `newFields` record column.

---

```powerquery
ExpandFields = Table.ExpandRecordColumn(
    ExpandSightings,
    "newFields",
    {"pct"},
    {"Location %"}
)
```

`Table.ExpandRecordColumn` extracts the `pct` field from the `newFields` record and renames it to `Location %`.

The `total` field is not expanded because it is not needed in the final output.

---

## Detailed example

For the location **Hales Wood**, the source data contains:

| Location | Species | No. of Sightings |
|---|---|---:|
| `Hales Wood` | `Grey` | `5` |
| `Hales Wood` | `Red` | `56` |

### Solution A walkthrough

After grouping with All Rows:

| Location | Details | Group Total |
|---|---|---:|
| `Hales Wood` | `Table (2 rows)` | `61` |

After expanding Details:

| Location | Group Total | Species | No. of Sightings |
|---|---:|---|---:|
| `Hales Wood` | `61` | `Grey` | `5` |
| `Hales Wood` | `61` | `Red` | `56` |

After calculating the percentage:

| Location | Species | No. of Sightings | Location % |
|---|---|---:|---:|
| `Hales Wood` | `Grey` | `5` | `8.20%` |
| `Hales Wood` | `Red` | `56` | `91.80%` |

### Solution B walkthrough

Inside `Table.Group`, the inner table for Hales Wood is:

| Location | Species | No. of Sightings |
|---|---|---:|
| `Hales Wood` | `Grey` | `5` |
| `Hales Wood` | `Red` | `56` |

`Table.AddColumn` processes each row:

For the Grey row (`n` = Grey row):

```text
total = List.Sum({5, 56}) = 61       ← from inner table column
pct = Value.Divide(5, 61) = 0.0820   ← from the current row; n[No. of Sightings] = 5
```

For the Red row (`n` = Red row):

```text
total = List.Sum({5, 56}) = 61       ← from inner table column
pct = Value.Divide(56, 61) = 0.9180  ← from the current row; n[No. of Sightings] = 56
```

The `newFields` column now contains a record for each row:

| Species | No. of Sightings | newFields |
|---|---:|---|
| `Grey` | `5` | `[total=61, pct=0.0820]` |
| `Red` | `56` | `[total=61, pct=0.9180]` |

After expanding `pct` as `Location %`, the result matches Solution A.

## Why this works

The challenge requires a calculation where the denominator is not a single value from the current row but a value derived from a group of rows. This is sometimes called a "window calculation".

Power Query does not have built-in window functions. Instead, `Table.Group` provides the mechanism for working with subsets of data.

**Solution A** is the approach that can be built entirely through the Power Query interface:

1. Use Group By, Advanced to create both a group total and an All Rows aggregation.
2. Expand the nested table to restore the original detail rows.
3. Calculate the percentage using the group total that now sits on each row.
4. Remove the Group Total helper column.

**Solution B** consolidates the work by performing the calculation inside the grouping step:

1. Use `Table.AddColumn` inside the `Table.Group` aggregation.
2. Use unique parameter names to distinguish row-level values from group-level values.
3. Expand the result.

Both solutions produce the same output. Solution A is more approachable and can be created through the UI. Solution B is more compact but requires understanding scope in nested tables.

## Notes

This solution assumes:

- the raw data query is named `RAW_DATA`
- the grouping column is `Location`
- the value column is `No. of Sightings`
- the percentage is calculated as each row's sighting count divided by the location total

If your raw data query has a different name, update this part:

```powerquery
Source = RAW_DATA
```

If your columns have different names, update the column references in the `Table.Group` and `Table.AddColumn` steps.

## Tags

`Power Query` · `M code` · `Table.Group` · `All Rows` · `Table.ExpandTableColumn` · `Table.AddColumn` · `List.Sum` · `Value.Divide` · `Scope` · `Inner table` · `Percentage` · `Excel` · `Power BI`
