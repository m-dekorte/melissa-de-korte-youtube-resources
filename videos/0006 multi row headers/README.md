# Safely Promote Multi-Row Headers

## Overview

This function promotes headers that span multiple rows in a table into a single header row.

In real-world data, column names sometimes appear across two or more rows. For example, a sales export might arrive with headers spread across three rows:

|  |  |  |  |  | Net | Unit | Gross |  |
|---|---|---|---|---|---|---|---|---|
|  | Item |  | Order |  | sales | cost | profit |  |
| Customer | number | Product | date | Quantity | amount | amount | amount | Status |

Where the intended column names are `Customer`, `Item number`, `Product`, `Order date`, `Quantity`, `Net sales amount`, `Unit cost amount`, `Gross profit amount`, and `Status`.

This function combines those rows into one header row, safely handling blanks, non-text values, missing headers, and duplicates along the way.

## Power Query M solution

```powerquery
(t as table, startRow as number, rowCount as number) as table =>
let
    Offset = startRow-1,
    Headers = List.Transform(Table.ToColumns(Table.Range(t, Offset, rowCount)), each
        let x = List.RemoveMatchingItems(_, {"", null}) in
        if List.IsEmpty(x) then null else Text.Combine(List.Transform(x, Text.From), " ")
    ),
    Transform = Table.PromoteHeaders(
        Table.InsertRows(Table.Skip(t, Offset+rowCount), 0, {Record.FromList(Headers, Table.ColumnNames(t))})
    )
in
    Transform
```

## How to use the solution

The solution is written as a custom function with three parameters:

| Parameter | Description |
|---|---|
| `t` | the source table |
| `startRow` | the row number where the header starts, using one-based counting |
| `rowCount` | the number of rows that make up the header |

For example, if the function is named `fxPromoteMultiRowHeaders`, and headers start on row 6 and span 3 rows:

```powerquery
fxPromoteMultiRowHeaders(Source, 6, 3)
```

The function removes everything above and including the header rows, inserts and then promotes the combined values as column names. The result is a table that starts directly at the data rows, with clean headers.

## Breakdown

The solution extracts the header rows, combines each column's values into a single header name, inserts those values on the first row then promotes that combined row as the new column headers.

---

```powerquery
(t as table, startRow as number, rowCount as number) as table =>
```

This defines a custom function that accepts a table `t`, a one-based starting row number `startRow`, and the number of rows that form the header `rowCount`.

The function returns a table.

---

```powerquery
Offset = startRow-1
```

This converts the one-based `startRow` to a zero-based offset.

Power Query uses zero-based indexing for row operations like `Table.Range` and `Table.Skip`, but users typically think of the first row as row 1 — that matches the row count they can see in the Preview Pane.

For example:

| `startRow` | `Offset` |
|---:|---:|
| `1` | `0` |
| `6` | `5` |
| `10` | `9` |

---

```powerquery
Table.Range(t, Offset, rowCount)
```

This extracts rows that contain header values.

`Table.Range` returns a subset of the table starting at the zero-based `Offset` position, returning `rowCount` rows.

For example, if headers start on row 6 and span 3 rows, `Table.Range` returns the content of rows 6, 7, and 8 as a three-row table.

---

```powerquery
Table.ToColumns(Table.Range(t, Offset, rowCount))
```

`Table.ToColumns` transposes the extracted header rows into a list of lists, where each inner list contains the values from one column across all header rows.

For example, given only these three header rows:

| Column1 | Column2 | Column3 |
|---|---|---|
|  |  | Net |
|  | Order | sales |
| Product | date | amount |

`Table.ToColumns` produces:

| Column | Values |
|---|---|
| `Column1` | `{null, null, "Product"}` |
| `Column2` | `{null, "Order", "date"}` |
| `Column3` | `{"Net", "sales", "amount"}` |

This is the key structural change. It groups values by column rather than by row, so they can be combined into a single header name per column.

---

```powerquery
List.Transform(Table.ToColumns(Table.Range(t, Offset, rowCount)), each
    let x = List.RemoveMatchingItems(_, {"", null}) in
    if List.IsEmpty(x) then null else Text.Combine(List.Transform(x, Text.From), " ")
)
```

`List.Transform` processes each column's list of header values.

For each column:

1. `List.RemoveMatchingItems(_, {"", null})` removes any blank or null values.
2. If all values were blank or null, the list is empty and the result is `null`.
3. Otherwise, `List.Transform(x, Text.From)` converts each remaining value to text, and `Text.Combine` joins them with a space.

The `Text.From` step is important because header row values are not always text. A cell might contain a number, a date, or another scalar type. `Text.Combine` expects text values, so `Text.From` ensures every value is converted to text before combining.

This also removes the need for `[PromoteAllScalars=true]` on `Table.PromoteHeaders`. By default, `Table.PromoteHeaders` only promotes text and number values to headers. Values of other scalar types, such as dates or logicals, would not be promoted and would receive a default column name instead. The `[PromoteAllScalars=true]` option extends promotion to all scalar types. Because this function converts everything to text with `Text.From` before promotion, the promoted row only ever contains text values or nulls, so the default behavior of `Table.PromoteHeaders` is sufficient.

Using the earlier example:

| Column | Values after cleanup | Combined header |
|---|---|---|
| `Column1` | `{"Product"}` | `Product` |
| `Column2` | `{"Order", "date"}` | `Order date` |
| `Column3` | `{"Net", "sales", "amount"}` | `Net sales amount` |

The nulls in `Column1` rows 1 and 2 are removed before combining, so the result is `Product` — not ` Product` with leading spaces.

If an entire column has no header values, the result is `null` instead of an empty string. This is deliberate. When `null` is later passed into the promoted header row, `Table.PromoteHeaders` falls back to its default column naming. This avoids creating columns with empty names.

---

```powerquery
Table.Skip(t, Offset+rowCount)
```

This removes the header rows and any rows above them from the original table.

The number of rows skipped is `Offset + rowCount`, which accounts for both:

- any rows before the header (`Offset`)
- the header rows themselves (`rowCount`)

For example, with `startRow = 6` and `rowCount = 3`, the offset is `5`, so `5 + 3 = 8` rows are skipped. This removes the report header, blank rows, and table header rows, leaving only rows from the table body onward.

---

```powerquery
Record.FromList(Headers, Table.ColumnNames(t))
```

This creates a record from the combined header values.

`Record.FromList` pairs each value in `Headers` with the corresponding column name from the original table. The original column names act as field names in the record, which allows it to be inserted into the table as a row.

---

```powerquery
Table.InsertRows(Table.Skip(t, Offset+rowCount), 0, {Record.FromList(Headers, Table.ColumnNames(t))})
```

`Table.InsertRows` inserts the combined header record as a new first row (position `0`) at the top of the data-only table.

The record is wrapped in curly braces `{ }` because `Table.InsertRows` expects a list of records, even when inserting a single row.

---

```powerquery
Table.PromoteHeaders(...)
```

`Table.PromoteHeaders` promotes the first row to become the column headers.

If any value in the promoted row is `null`, `Table.PromoteHeaders` assigns a default name based on the column's position in the table, such as `Column1`, `Column6`, or `Column12`. The numbering reflects each column's position, not a count of how many columns needed a default name.

`Table.PromoteHeaders` also enforces unique column names. If the combined headers produce duplicate names, Power Query disambiguates them by appending a numeric suffix preceded by a dot. For example, two columns both named `amount` would become `amount` and `amount.1`.

---

## Detailed example

The sample data for this function represents a sales export. The worksheet contains 12 columns. Rows 1 through 3 hold metadata, rows 4 and 5 are blank, rows 6 through 8 contain the column headers, and rows 9 onward contain data.

The three header rows look like this:

| Col1 | Col2 | Col3 | Col4 | Col5 | Col6 | Col7 | Col8 | Col9 | Col10 | Col11 | Col12 |
|---|---|---|---|---|---|---|---|---|---|---|---|
|  |  |  |  |  |  |  | Net | Unit | Gross |  |  |
|  |  |  | Item |  | Order |  | sales | cost | profit |  |  |
|  |  | Customer | number | Product | date | Quantity | amount | amount | amount | Status |  |

This structure includes:

- columns with no header values at all (`Col1`, `Col2`, `Col12`)
- single-row headers (`Customer`, `Product`, `Quantity`, `Status`)
- two-row headers (`Item number`, `Order date`)
- three-row headers (`Net sales amount`, `Unit cost amount`, `Gross profit amount`)

Calling the function with `startRow = 6` and `rowCount = 3`:

**Step 1:** `Offset` becomes `5`.

**Step 2:** `Table.Range` extracts the three header rows (rows 6, 7, and 8).

**Step 3:** `Table.ToColumns` transposes to column-based lists:

| Column | Row 6 | Row 7 | Row 8 |
|---|---|---|---|
| `Col1` |  |  |  |
| `Col2` |  |  |  |
| `Col3` |  |  | Customer |
| `Col4` |  | Item | number |
| `Col5` |  |  | Product |
| `Col6` |  | Order | date |
| `Col7` |  |  | Quantity |
| `Col8` | Net | sales | amount |
| `Col9` | Unit | cost | amount |
| `Col10` | Gross | profit | amount |
| `Col11` |  |  | Status |
| `Col12` |  |  |  |

**Step 4:** Blanks and nulls are removed, remaining values are converted to text and combined:

| Column | After cleanup | Combined |
|---|---|---|
| `Col1` | `{}` | `null` |
| `Col2` | `{}` | `null` |
| `Col3` | `{"Customer"}` | `Customer` |
| `Col4` | `{"Item", "number"}` | `Item number` |
| `Col5` | `{"Product"}` | `Product` |
| `Col6` | `{"Order", "date"}` | `Order date` |
| `Col7` | `{"Quantity"}` | `Quantity` |
| `Col8` | `{"Net", "sales", "amount"}` | `Net sales amount` |
| `Col9` | `{"Unit", "cost", "amount"}` | `Unit cost amount` |
| `Col10` | `{"Gross", "profit", "amount"}` | `Gross profit amount` |
| `Col11` | `{"Status"}` | `Status` |
| `Col12` | `{}` | `null` |

**Step 5:** `Table.Skip` removes the first 8 rows (5 rows above the header plus 3 header rows), leaving only data. The combined header values are inserted as the first row of that data-only table.

**Step 6:** `Table.PromoteHeaders` promotes the first row to column headers. The `null` values in columns 1, 2, and 12 receive default names based on their position:

| Column1 | Column2 | Customer | Item number | Product | Order date | Quantity | Net sales amount | Unit cost amount | Gross profit amount | Status | Column12 |
|---|---|---|---|---|---|---|---|---|---|---|---|
|  |  | Acme Retail | ITM-1001 | Wireless Mouse | 6/1/2026 | 12 | 299.40 | 12.00 | 155.40 | Shipped |  |
|  |  | Contoso Stores | ITM-1002 | Mechanical Keyboard | 6/2/2026 | 5 | 395.00 | 46.00 | 165.00 | Shipped |  |
|  |  | Fabrikam Office | ITM-1003 | USB-C Dock | 6/3/2026 | 3 | 388.50 | 82.00 | 142.50 | Pending |  |

The default names are `Column1`, `Column2`, and `Column12`. `Table.PromoteHeaders` assigns default names based on each column's position in the table, not by counting how many columns needed a default name.

The empty `Column1`, `Column2`, and `Column12` columns can be removed afterward with `Table.RemoveColumns` or by selecting the columns you want to keep with `Table.SelectColumns`.

## Why this works

The key idea is to treat multiline headers as a transposition problem.

Instead of trying to merge rows horizontally, the function transposes the header rows into column-based lists using `Table.ToColumns`. This makes it easy to combine values per column with `Text.Combine`, while safely ignoring blanks with `List.RemoveMatchingItems`.

The approach is safe because:

- blank or null cells in the header rows do not produce leading or trailing spaces
- non-text values such as numbers or dates are converted to text by `Text.From` before combining, so `Table.PromoteHeaders` always receives text values or nulls
- columns with no header values at all become `null`, which causes `Table.PromoteHeaders` to assign a default column name based on position rather than creating a column with an empty name
- duplicate header names are automatically disambiguated by `Table.PromoteHeaders` with a numeric suffix
- the original data rows are preserved without modification

## Notes

### Assumptions

This function assumes:

- the header rows are contiguous
- `startRow` uses one-based counting — row 1 is the first row visible in the Preview Pane
- blank cells in the header rows should be ignored, not treated as part of the header name
- the combined header values are joined with a single space

### Using the function in your own queries

To use this function, create a new blank query, open the Advanced Editor, and paste in the function code. Name the query something recognizable, such as `fxPromoteMultiRowHeaders`.

Then invoke it from any other query by calling:

```powerquery
fxPromoteMultiRowHeaders(PreviousStepName, 6, 3)
```

Replace `PreviousStepName` with the name of the step that produces the table you want to transform. Replace `6` and `3` with the actual start row and row count for your data.

If the header is only one row deep, `rowCount = 1` still works. The function extracts that single row, removes blanks, and promotes it, which is equivalent to a standard `Table.PromoteHeaders` but with null safety.

If the header rows start at the very top of the table, use `startRow = 1`.

### Troubleshooting

**Type mismatch error when inspecting intermediate steps**

While building or debugging the function, you may want to return an intermediate value — such as the `Headers` list — to inspect it in the Preview Pane. If the function's return type is set to `as table` and the intermediate value is not a table, Power Query raises a type mismatch error.

To work around this, temporarily comment out the return type:

```powerquery
(t as table, startRow as number, rowCount as number) /* as table */ =>
```

This removes the type constraint so you can return any value. Restore it when you are done inspecting.

**Off-by-one errors with startRow**

The `startRow` parameter is one-based, matching what you see in the Preview Pane. If you pass a zero-based index by mistake, the function will extract the wrong rows. If your headers start on the row labeled `6` in the Preview Pane, pass `6`.

**Columns you do not need**

After promoting headers, you may have columns with default names like `Column1` or `Column12` that hold no useful data. Remove them with a follow-up step:

```powerquery
Table.RemoveColumns(fxPromoteMultiRowHeaders(Source, 6, 3), {"Column1", "Column2", "Column12"})
```

Or select only the columns you want to keep:

```powerquery
Table.SelectColumns(fxPromoteMultiRowHeaders(Source, 6, 3), {"Customer", "Item number", "Product", "Order date", "Quantity", "Net sales amount", "Unit cost amount", "Gross profit amount", "Status"})
```

## Tags

`Power Query` · `M code` · `Custom function` · `Table.PromoteHeaders` · `Table.Range` · `Table.ToColumns` · `Table.InsertRows` · `Record.FromList` · `List.RemoveMatchingItems` · `Text.From` · `Text.Combine` · `Excel` · `Power BI`
