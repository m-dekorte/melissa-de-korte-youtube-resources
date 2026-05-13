# Challenge 399: Table Transformation

## Overview

This challenge is about transforming a column of small, text-based table structures into one combined result table.

Each value in the source column contains:

- a header line
- a data line

The number and order of fields may differ per row, so the solution needs to read the headers dynamically and align the results into one table.

## Challenge rules

Transform the question structure into the result structure.

The solution should be dynamic, meaning it should not rely on hard-coded column positions for every possible structure.

## Source

Challenge creator: [Omid Motamedisedeh on LinkedIn](https://www.linkedin.com/in/omidmot/)

Challenge file: [Download Excel file](https://lnkd.in/g9vvMs_Q)

## Sample data

The input contains one column, where each cell contains a two-line structure.

Tabs are shown as `↹` below for readability.

| Column1 |
|---|
| `Date↹Customer↹Product↹Sale`<br>`2/12/2024↹C1↹A↹10` |
| `Date↹Product↹Sale`<br>`2/12/2024↹B↹4` |
| `Date↹Customer↹Product↹Sale`<br>`2/12/2024↹C2↹A↹2` |
| `Customer↹Product↹Sale`<br>`C1↹A↹1` |
| `Date↹Customer↹Product↹Sale`<br>`13/12/2024↹C1↹C↹11` |
| `Date↹Customer`<br>`14/12/2024↹C1` |

Expected result:

| Date | Customer | Product | Sale |
|---|---|---|---:|
| `2/12/2024` | `C1` | `A` | `10` |
| `2/12/2024` |  | `B` | `4` |
| `2/12/2024` | `C2` | `A` | `2` |
|  | `C1` | `A` | `1` |
| `13/12/2024` | `C1` | `C` | `11` |
| `14/12/2024` | `C1` |  |  |

## Power Query M solution

```powerquery
Table.Combine(
    List.Transform(
        Excel.CurrentWorkbook(){[Name = "Table1"]}[Content][Column1], 
        each [
            l = Lines.FromText(_), 
            a = Table.FromRows({Text.Split(l{1}, "#(tab)")}, Text.Split(l{0}, "#(tab)"))
        ][a]
    )
)
```

## Breakdown

The solution uses `List.Transform` to turn each text value into a small table, and then uses `Table.Combine` to combine those tables into one result.

---

```powerquery
Excel.CurrentWorkbook(){[Name = "Table1"]}[Content][Column1]
```

This gets the `Column1` values from the Excel table named `Table1`.

Instead of returning the full table, `[Column1]` extracts the values from that column as a list.

Each item in this list is one two-line text structure from the question column.

---

```powerquery
List.Transform(
    Excel.CurrentWorkbook(){[Name = "Table1"]}[Content][Column1],
    each ...
)
```

`List.Transform` loops through each value in `Column1`.

For every item, it creates a separate one-row table based on the headers and values inside that text item.

---

```powerquery
Lines.FromText(_)
```

`Lines.FromText` splits the current text value into separate lines.

For each source value:

- `l{0}` is the first line, containing the column names
- `l{1}` is the second line, containing the row values

Power Query uses zero-based indexing, so the first item in a list is item `0`.

Example:

```text
Date    Customer    Product    Sale
2/12/2024    C1    A    10
```

becomes a list like this:

| List item | Meaning |
|---:|---|
| `l{0}` | header line |
| `l{1}` | value line |

---

```powerquery
Text.Split(l{0}, "#(tab)")
```

This splits the header line by tab characters.

For example:

```text
Date    Customer    Product    Sale
```

becomes:

```text
Date, Customer, Product, Sale
```

These values become the column names for the small table.

---

```powerquery
Text.Split(l{1}, "#(tab)")
```

This splits the value line by tab characters.

For example:

```text
2/12/2024    C1    A    10
```

becomes:

```text
2/12/2024, C1, A, 10
```

These values become the data row for the small table.

---

```powerquery
Table.FromRows(
    {Text.Split(l{1}, "#(tab)")},
    Text.Split(l{0}, "#(tab)")
)
```

`Table.FromRows` creates a table from the split values.

The first argument provides the row values.

The second argument provides the column names.

The value line is wrapped in `{ }` because `Table.FromRows` expects a list of rows, even when there is only one row.

For example, this source value:

```text
Date    Product    Sale
2/12/2024    B    4
```

creates a one-row table like this:

| Date | Product | Sale |
|---|---|---:|
| `2/12/2024` | `B` | `4` |

---

```powerquery
[
    l = Lines.FromText(_),
    a = Table.FromRows({Text.Split(l{1}, "#(tab)")}, Text.Split(l{0}, "#(tab)"))
][a]
```

This part creates a temporary record.

The record has two fields:

| Field | Purpose |
|---|---|
| `l` | stores the split lines from the current text value |
| `a` | creates the one-row table from those lines |

The final `[a]` returns only the table stored in the `a` field.

This is a compact way to store an intermediate result without writing a separate `let ... in` expression.

---

```powerquery
Table.Combine(...)
```

`Table.Combine` combines all the one-row tables into one final table.

This is what makes the solution dynamic.

Because each small table is created from its own header line, rows do not need to contain all columns. When a row is missing a column, Power Query leaves that value blank/null in the combined result.

For example:

| Source structure | Columns created |
|---|---|
| `Date Customer Product Sale` | Date, Customer, Product, Sale |
| `Date Product Sale` | Date, Product, Sale |
| `Customer Product Sale` | Customer, Product, Sale |
| `Date Customer` | Date, Customer |

After combining, Power Query aligns the matching column names into one result table.

## Why this works

The key idea is to treat each source value as a tiny table.

Each item contains its own headers and values, so the solution:

1. splits the text into lines
2. uses the first line as column names
3. uses the second line as row values
4. creates a one-row table
5. combines all one-row tables into one final table

This avoids hard-coding the possible structures.

The result adapts to rows with missing fields, different field combinations, or a different order of fields, as long as each source value follows the same two-line, tab-separated pattern.

## Notes

This solution assumes:

- the Excel table is named `Table1`
- the source column is named `Column1`
- each cell contains two lines
- the first line contains tab-separated column names
- the second line contains tab-separated values

If your table has a different name, update this part:

```powerquery
Excel.CurrentWorkbook(){[Name = "Table1"]}[Content]
```

If your source column has a different name, update this part:

```powerquery
[Column1]
```

## Tags

`Power Query` · `M code` · `Table transformation` · `Table.Combine` · `Table.FromRows` · `Lines.FromText` · `Text.Split` · `Excel`
