# Beware the Scan Limit — Potential Data Loss in Power Query

Power Query uses sampling and scan row limits. Some limits are documented, some are visible in the user interface, and others are easy to miss. This resource lists limits covered in the video, along with sample M code to reproduce each scenario.

Video: [Watch on YouTube](https://youtu.be/1VpKlmhEZDo)

LinkedIn: [Connect with Melissa](https://www.linkedin.com/in/melissa-de-korte/)

---

## Quick Reference

| # | Feature | Scan Limit | Warning? | Documented? | Risk |
|--:|---|--:|---|---|---|
| 1 | Data Profiling | 1000 rows | Yes | Yes | Low |
| 2 | Nested Table or List structure (secondary preview) | 20 items | No | No | Low |
| 3 | Auto Filter list | 1000 distinct values | Yes | Yes | Low |
| 4 | Detect Data Type | 200 rows | No | Yes | High |
| 5 | Column From Examples | 100 rows | No | Yes | Medium |
| 6 | Expand Table/Record fields | 1000 rows | Yes | No | Medium |
| 7 | Split Column | 200 rows | No | No | High |

Checked August 2026

---

## Limits Covered in the Video

### 1. Data Profiling — Top 1,000 Rows

By default, **Column Profiling** (column quality, column profile, column distribution) is based on the first **1,000 rows**. If data quality issues exist beyond that range, profiling won't report them.

The status bar shows: *Column profiling based on top 1000 rows.* Click it to switch to *entire data set* when validating data. Switch back when done — profiling the entire dataset has a performance cost.

```m
// DataProfiling
let
    Source = Table.FromColumns({{1..1000}})
        & Table.FromColumns({List.Repeat({null}, 1000)})
in
    Source
```

The first 1,000 rows contain values. The next 1,000 contain nulls.  
With default profiling, Power Query reports zero percent empty.


>Documentation  
[Data Profiling Tools](https://learn.microsoft.com/en-us/power-query/data-profiling-tools?wt.mc_id=MVP_371969)

---

### 2. Nested Structures — Secondary Preview Top 20

The **secondary preview** (that appears when you click the whitespace next to a nested structure) shows at most **20 items** for nested tables and nested lists. 
Records are not limited in the secondary preview.

```m
// NestedStructures
let
    Source = Table.FromValue([table = Table.FromColumns({{1..50}})])
        & Table.FromValue([list = {1..50}])
        & Table.FromValue(
            [record = Record.FromList(
                {1..10001},
                List.Transform({1..10001}, Text.From)
            )]
        )
in
    Source
```

The nested table and list each contain 50 items but the secondary preview only shows 20.  
The record contains 10,001 fields and all are visible.


>Documentation  
VIEWABLE RANGE IS UNDOCUMENTED  

---

### 3. Filter Values — Top 1,000 Distinct Values

The **Auto Filter list** loads at most **1,000 distinct values**. Power Query can warn you the list may be incomplete and may offer to *Load more*. Even after loading more, the list may still be incomplete.

```m
// FilterValues
let
    Source = Table.FromColumns({{1..1001}})
in
    Source
```

One column with 1,001 unique values.  
The filter list loads 1,000 of them.


>Documentation  
[Auto Filter list](https://learn.microsoft.com/en-us/power-query/filter-values?wt.mc_id=MVP_371969)

---

### 4. Type Detection — Top 200 Rows

**Detect Data Type** as well as **Enabled Automatic Type Detection** will inspect the first **200 rows** to infer column types. (This applies to unstructured sources such as Excel, CSV, and text files.) The limit is documented.

The dangerous part: assigning a wrong type does not have to raise an error. A decimal value will be truncated to a whole number without warning.

```m
// TypeDetection
let
    Source = Table.Repeat(Table.FromColumns({{15}}), 200)
        & Table.FromColumns({{15.99}})
in
    Source
```

The first 200 rows contain the whole number `15`. Row 201 contains `15.99`. Using the option **Detect Data Type** assigns `Int64.Type`, the value `15.99` becomes `15` — no error, no warning.


>Documentation  
[Data Type Detection](https://learn.microsoft.com/en-us/power-query/data-types?wt.mc_id=MVP_371969)

---

### 5. Column From Examples — Top 100 Rows

The **Column From Examples** experience works with only the top **100 rows** of the data preview. This is documented. If a pattern variation appears first beyond row 100, it is not available in the Column From Examples interface.

```m
// ColFromExample
let
    Source = Table.FromColumns({{1..1001}})
in
    Source
```

**Workaround:** Create a temporary step that places representative rows (covering all pattern variations) in the top 100, generate the expression, then remove the temporary step.


>Documentation  
[Column from Examples - tips and considerations](https://learn.microsoft.com/en-us/power-query/column-from-example?wt.mc_id=MVP_371969)

---

## The Exception: List Expansion

Expanding a column of lists does **not** require scanning or sampling. Both *Expand to New Rows* and *Extract Values* process every item regardless of position.

```m
// ExpandLst
let
    Source = Table.Repeat(Table.FromValue([A = {1..3}]), 1000)
        & Table.FromValue([A = {1..4}]),
    ExpandToNewRows = Table.ExpandListColumn(Source, "Value"),
    ExtractValues = Table.TransformColumns(
        Source,
        {"Value", each Text.Combine(List.Transform(_, Text.From), ", "), type text}
    )
in
    ExtractValues
```

The first 1,000 rows contain 3-item lists. Row 1,001 contains a 4-item list.  
Both expand operations handle all items correctly.

---

### 6. Table/Record Expansion — Top 1,000 Row Discovery Limit

When **expanding a column** of nested tables or records, Power Query inspects the first **1,000 rows** to discover available field names. If a field appears first beyond row 1,000, it won't be listed in the expand dialog.

The UI may indicate the scan limit was reached. Unmatched column names in the `columnNames` list don't raise errors.

**Nested tables:**

```m
// ExpandTblToCols
let
    Source = Table.Repeat(
            Table.FromValue([A = Table.FromRecords({[1 = 1, 2 = 2]})]),
            1000
        )
        & Table.FromValue([A = Table.FromRecords({[1 = 1, 2 = 2, 3 = 3]})])
in
    Source
```

**Nested records:**

```m
// ExpandRecToCols
let
    Source = Table.Repeat(
            Table.FromValue([A = [1 = 1, 2 = 2]]),
            1000
        )
        & Table.FromValue([A = [1 = 1, 2 = 2, 3 = 3]])
in
    Source
```

The first 1,000 rows contain two fields. Row 1,001 contains a third field that won't appear in the expand dialog.

**Workaround:** Add a temporary column that collects all nested column names (e.g. using `Table.ColumnNames` or `Record.FieldNames`), inspect the variations, and manually edit the `columnNames` list in the generated code.


>Documentation  
SCAN RANGE IS UNDOCUMENTED  
[Expand Table Column](https://learn.microsoft.com/en-us/powerquery-m/table-expandtablecolumn?wt.mc_id=MVP_371969)  
[Expand Record Column](https://learn.microsoft.com/en-us/powerquery-m/table-expandrecordcolumn?wt.mc_id=MVP_371969)

---

### 7. Split Column — Top 200 Rows

A **Split Column** operation inspects the first **200 rows** to determine how many destination columns are needed. This limit is not documented. Unlike the expand operations, there is no warning that a scan limit was reached.

The generated `Table.SplitColumn` function uses `ExtraValues.Ignore` by default. When a split produces more values than there are destination columns, the extra values are dropped.

```m
// SplitToCols
let
    Source = Table.Repeat(Table.FromValue("1, 2, 3"), 200)
        & Table.FromValue("1, 2, 3, 4")
in
    Source
```

The first 200 rows split into 3 values. Row 201 splits into 4 values. The dialog creates 3 destination columns. The fourth value is ignored without warning.

**Workaround:** After splitting, set the optional `extraColumns` parameter to `ExtraValues.Error` or `ExtraValues.List` to trigger an error or surface any extra values instead of discarding them.

>Documentation  
SCAN RANGE IS UNDOCUMENTED  
[Split Column](https://learn.microsoft.com/en-us/powerquery-m/table-splitcolumn?wt.mc_id=MVP_371969)

---

## Other Limits in Power Query

### 8. Enter Data — 3,000 Cells

The **Enter Data** feature stores the values you enter in a compressed binary format. It supports a maximum of **3,000 populated cells**, which you can paste directly into the Enter Data window.

### 9. Cell String Length — 32,766 Characters

The maximum length of a text string in a single cell is **32,766 characters**. Values that exceed this length (such as image binary strings or long text blocks) will cause failures unless they are split across multiple rows.

---

## Key Takeaway

The problem isn't that sampling exists. Without limits and previews, Power Query would need to inspect far more data every time you interact with the editor — slowing down the entire query development process.

The problem is assuming a sample represents the entire dataset.

Being aware of these limits helps you design better queries and avoid potential data loss.

---

## Tags

`Power Query` · `M code` · `Scan limits` · `Data profiling` · `Type detection` · `Split column` · `Expand` · `Data loss` · `Excel` · `Power BI`
