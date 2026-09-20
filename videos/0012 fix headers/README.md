# Fix Missing and Shifting Headers with M

This challenge is about cleaning up column headers in a table exported from an external system.

The solution needs to reformat the month headers, insert the missing description header, promote the cleaned values to proper column names and should be dynamic.

- Video: [Watch on YouTube](https://youtu.be/hCI1ui-dGQQ)
- LinkedIn: [Connect with Melissa](https://www.linkedin.com/in/melissa-de-korte/)
- Challenge creator: [Shirley Moreman](https://www.linkedin.com/in/shirleymoreman/)

## Scenario

The source table arrives with auto-generated column names. The first row contains the actual headers in their raw format.

## Source data

Source table:

| Column1 | Column2 | Column3 | Column4 | Column5 | Column6 | Column7 | Column8 | Column9 | Column10 | Column11 |
|---|---|---|---|---|---|---|---|---|---|---|
| `Stock Code` | ` 4 / 2025` | ` 5 / 2025` | ` 6 / 2025` | ` 7 / 2025` | ` 8 / 2025` | ` 9 / 2025` | `10 / 2025` | `11 / 2025` | `12 / 2025` | `null` |
| `A113` | `DRILL BIT` | `5` | `3` | `7` | `2` | `6` | `5` | `10` | `2` | `1` |
| `Last Cost :` | `2.2248` | `Current In Stock :` | `5` | | | | | | | |
| `A223` | `DRILL BIT` | `3` | `6` | `8` | `2` | `1` | `2` | `5` | `3` | `0` |
| `Last Cost :` | `2.2248` | `Current In Stock :` | `3` | | | | | | | |
| `A456` | `HAMMER` | `2` | `5` | `1` | `3` | `2` | `1` | `2` | `0` | `1` |
| `Last Cost :` | `12.3291` | `Current In Stock :` | `1` | | | | | | | |

## Challenge rules

Fix the column headers using these rules:

1. The first row of the table contains the actual column headers.
2. The month values in the header row are formatted as `" m / yyyy"` (with leading spaces) and must be converted to `"yyyy-MM"`.
3. A `Description` header is missing between the first column and the month columns. It must be inserted at position 2 (the second column).
4. The solution should be dynamic, meaning it should work regardless of which nine months appear in the file.

Expected result:

| Stock Code | Description | 2025-04 | 2025-05 | 2025-06 | 2025-07 | 2025-08 | 2025-09 | 2025-10 | 2025-11 | 2025-12 |
|---|---|---|---|---|---|---|---|---|---|---|
| `A113` | `DRILL BIT` | `5` | `3` | `7` | `2` | `6` | `5` | `10` | `2` | `1` |
| `Last Cost :` | `2.2248` | `Current In Stock :` | `5` | | | | | | | |
| `A223` | `DRILL BIT` | `3` | `6` | `8` | `2` | `1` | `2` | `5` | `3` | `0` |
| `Last Cost :` | `2.2248` | `Current In Stock :` | `3` | | | | | | | |
| `A456` | `HAMMER` | `2` | `5` | `1` | `3` | `2` | `1` | `2` | `0` | `1` |
| `Last Cost :` | `12.3291` | `Current In Stock :` | `1` | | | | | | | |

The column headers have been cleaned:

- `" 4 / 2025"` became `2025-04`
- the missing `Description` header was inserted
- the trailing null column was removed
- the original header row was removed from the data

## Raw data query

This creates the sample dataset. Name this query `RAW_DATA` so the solution can reference it.

```powerquery
let
    Source = Table.FromRows(
        {
            {"Stock Code", " 4 / 2025", " 5 / 2025", " 6 / 2025", " 7 / 2025", " 8 / 2025", " 9 / 2025", "10 / 2025", "11 / 2025", "12 / 2025", null},
            {"A113", "DRILL BIT", 5, 3, 7, 2, 6, 5, 10, 2, 1},
            {"Last Cost :", 2.2248, "Current In Stock :", 5, null, null, null, null, null, null, null},
            {"A223", "DRILL BIT", 3, 6, 8, 2, 1, 2, 5, 3, 0},
            {"Last Cost :", 2.2248, "Current In Stock :", 3, null, null, null, null, null, null, null},
            {"A456", "HAMMER", 2, 5, 1, 3, 2, 1, 2, 0, 1},
            {"Last Cost :", 12.3291, "Current In Stock :", 1, null, null, null, null, null, null, null},
            {"A567", "SCREWDRIVER", 3, 2, 5, 9, 8, 6, 0, 0, 1},
            {"Last Cost :", 13.204858, "Current In Stock :", 2, null, null, null, null, null, null, null},
            {"A951", "NAIL GUN", 4, 2, 1, 0, 0, 0, 1, 0, 1},
            {"Last Cost :", 13.204858, "Current In Stock :", 1, null, null, null, null, null, null, null},
            {"A326", "NAIL REFILL PACK", 3, 0, 0, 0, 0, 2, 0, 0, 2},
            {"Last Cost :", 2.7589, "Current In Stock :", 1, null, null, null, null, null, null, null}
        }
    )
in
    Source
```
---

## Power Query M solution

```powerquery
let
    Source = RAW_DATA,
    Transform = List.Transform(
        Record.ToList(Source{0}),
        each try Date.ToText(Date.From(_, ""), "yyyy-MM") otherwise _
    ),
    fixedHeaders = List.InsertRange(
        List.RemoveNulls(Transform),
        1,
        {"Description"}
    ),
    renamedCol = Table.Skip(
        Table.RenameColumns(Source,
            List.Zip({Table.ColumnNames(Source), fixedHeaders})
        )
    )
in
    renamedCol
```

## Breakdown

The solution reads the first row to extract the raw headers, cleans them up, inserts the missing header, renames the columns, and removes the original header row.

---

```powerquery
Source = RAW_DATA
```

This references the query named `RAW_DATA`.

Because the data comes from an external system, the table has auto-generated column names like `Column1`, `Column2`, and so on. The actual headers are stored in the first row of data.

---

```powerquery
Record.ToList(Source{0})
```

`Source{0}` gets the first row of the table as a record.

`Record.ToList` converts that record into a list of values.

For example, the first row:

| Column1 | Column2 | Column3 | ... | Column10 | Column11 |
|---|---|---|---|---|---|
| `Stock Code` | ` 4 / 2025` | ` 5 / 2025` | ... | `12 / 2025` | `null` |

becomes the list:

```powerquery
{"Stock Code", " 4 / 2025", " 5 / 2025", ..., "12 / 2025", null}
```

---

```powerquery
Transform = List.Transform(
    Record.ToList(Source{0}),
    each try Date.ToText(Date.From(_, ""), "yyyy-MM") otherwise _
)
```

`List.Transform` processes each value from the first row.

For each value, the function tries to parse it as a date using `Date.From` and then format it as `"yyyy-MM"` using `Date.ToText`.  
If the value cannot be parsed as a date, the `try ... otherwise` expression returns the original value unchanged.

| Original value | Can parse as date? | Result |
|---|---|---|
| `Stock Code` | no | `Stock Code` |
| ` 4 / 2025` | yes (April 2025) | `2025-04` |
| ` 5 / 2025` | yes (May 2025) | `2025-05` |
| `12 / 2025` | yes (December 2025) | `2025-12` |
| `null` | no | `null` |

The second argument to `Date.From` is an empty string `""`, which specifies the invariant culture. This allows Power Query to interpret the month/year text values as dates regardless of the user's locale settings.

The function does not need to know which months appear in the file. Any text value that Power Query can interpret as a date is automatically reformatted.

---

```powerquery
List.RemoveNulls(Transform)
```

This removes null values from the transformed list.

The original header row contains a trailing null in the last column. After removing it, the list becomes:

```powerquery
{"Stock Code", "2025-04", "2025-05", ..., "2025-12"}
```

This list has 10 items: one stock code header and nine month headers.

---

```powerquery
fixedHeaders = List.InsertRange(
    List.RemoveNulls(Transform),
    1,
    {"Description"}
)
```

`List.InsertRange` inserts one or more items into a list at a specific position. Positions in M are zero-based.

Here it inserts `"Description"` at position `1`, which is after `"Stock Code"` and before the first month column.

Before:

```powerquery
{"Stock Code", "2025-04", "2025-05", ..., "2025-12"}
```

After:

```powerquery
{"Stock Code", "Description", "2025-04", "2025-05", ..., "2025-12"}
```

This list now has 11 items, matching the 11 columns in the source table.

The `Description` header was missing because the exported system does not include a header for the product description column. The description values (such as `DRILL BIT` or `HAMMER`) appear in the second column of each product row, but the header row skips that column entirely.

---

```powerquery
List.Zip({Table.ColumnNames(Source), fixedHeaders})
```

This pairs each original column name with its new name.

`Table.ColumnNames(Source)` returns the auto-generated names:

```powerquery
{"Column1", "Column2", "Column3", ..., "Column11"}
```

`fixedHeaders` contains the cleaned header names:

```powerquery
{"Stock Code", "Description", "2025-04", ..., "2025-12"}
```

`List.Zip` combines them into pairs:

| Original name | New name |
|---|---|
| `Column1` | `Stock Code` |
| `Column2` | `Description` |
| `Column3` | `2025-04` |
| `Column4` | `2025-05` |
| ... | ... |
| `Column11` | `2025-12` |

---

```powerquery
Table.RenameColumns(
    Source,
    List.Zip({Table.ColumnNames(Source), fixedHeaders})
)
```

`Table.RenameColumns` applies the rename pairs to the source table.

After this step, all columns have their correct names. However, the first row of data still contains the original raw header values.

---

```powerquery
renamedCol = Table.Skip(
    Table.RenameColumns(...)
)
```

`Table.Skip` removes the first row from the table.

That first row contained the original header values (such as `Stock Code`, ` 4 / 2025`, and so on), which are no longer needed because the columns have been renamed.

After skipping, the data starts with the first product row.

## Detailed example

Starting with the source table, the first row contains:

```text
Stock Code,  4 / 2025,  5 / 2025,  6 / 2025,  7 / 2025,  8 / 2025,  9 / 2025, 10 / 2025, 11 / 2025, 12 / 2025, null
```

Step 1: Each value is tested with `Date.From`. The month values parse as dates and are reformatted:

| Before | After |
|---|---|
| ` 4 / 2025` | `2025-04` |
| ` 5 / 2025` | `2025-05` |
| ` 9 / 2025` | `2025-09` |
| `10 / 2025` | `2025-10` |
| `Stock Code` | `Stock Code` |
| `null` | `null` |

Step 2: Nulls are removed, leaving 10 values.

Step 3: `"Description"` is inserted at (zero-based index) position 1:

```text
Stock Code, Description, 2025-04, 2025-05, ..., 2025-12
```

Step 4: The 11 auto-generated column names are paired with the 11 fixed headers and renamed.

Step 5: The first row is skipped, leaving the data rows with clean column names.

The result for the first product row:

| Stock Code | Description | 2025-04 | 2025-05 | 2025-06 | 2025-07 | 2025-08 | 2025-09 | 2025-10 | 2025-11 | 2025-12 |
|---|---|---|---|---|---|---|---|---|---|---|
| `A113` | `DRILL BIT` | `5` | `3` | `7` | `2` | `6` | `5` | `10` | `2` | `1` |

## Why this works

The solution treats the header cleanup as three independent problems:

1. **Reformatting month values** — Instead of parsing the month text with string operations, the solution uses `Date.From` to interpret each value as a date. This works because Power Query can recognize the `" m / yyyy"` pattern as a valid date. Once parsed, `Date.ToText` formats it cleanly as `"yyyy-MM"`. The `try ... otherwise` pattern means non-date values pass through unchanged, so only the month headers are affected.

2. **Inserting the missing header** — The `Description` header is missing from the exported data because the system does not include it. Since the description values always appear in the second column, `List.InsertRange` adds `"Description"` at position `1` to fill the gap. This shifts the month headers to align with their correct data columns.

3. **Removing the null column** — The trailing null in the header row corresponds to an empty column in the export. `List.RemoveNulls` removes it before the new header list is built, so the rename operation produces the correct column pairs.

The combination of `try ... otherwise` and `Date.From` makes this dynamic. The solution does not hard-code any specific month names or positions. If next month's export contains different months, the same code handles them automatically.

## Notes

This solution assumes:

- the source data query is named `RAW_DATA`
- the first row of data contains the raw column headers
- month values in the header row can be parsed as dates by Power Query
- the `Description` header is always the second column
- there are always nine month columns
- there is one trailing null column to remove

If your table has a different name, update this part:

```powerquery
Source = RAW_DATA
```

If the missing header has a different name or position, update this part:

```powerquery
List.InsertRange(List.RemoveNulls(Transform), 1, {"Description"})
```

This solution addresses the header cleanup only. The remaining data still contains the alternating product rows and `Last Cost` / `Current In Stock` rows from the original export. Further transformation steps would be needed to separate those into a final clean table.

## Tags

`Power Query` · `M code` · `Header cleanup` · `Table.RenameColumns` · `Table.Skip` · `Date.From` · `Date.ToText` · `List.InsertRange` · `List.RemoveNulls` · `Record.ToList` · `List.Zip` · `Excel` · `Power BI`
