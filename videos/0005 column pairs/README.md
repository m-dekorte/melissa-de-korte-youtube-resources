# Unpivot Multiple Column Pairs Dynamically in Power Query

A step-by-step guide to building a dynamic solution that unpivots multiple column pairs (or column sets) instead of hard-coding column names.

- **Video:** [Unpivot Multiple Column Pairs Dynamically in Power Query](https://youtu.be/sRJdVyDFWlY)
- **LinkedIn:** [Connect with Melissa](https://www.linkedin.com/in/melissa-de-korte-64585884)

---

## Scenario covered

HR exports often store overtime data in a pivoted layout — one pair of Hours and Cost columns per overtime rate. The number of rates can vary between exports.

This video builds a dynamic solution that reads the column structure from the data itself, so it continues to work when rates are added or removed — provided the naming convention stays consistent.

Two layouts handled:

- **Paired layout** — `OT Hours @x` and `OT Cost @x` appear side by side for each rate
- **Split layout** — all `OT Hours @x` columns appear first, followed by all `OT Cost @x` columns

---

## Files in this folder

| File | Description |
|---|---|
| `00_EXPECTED_OUTCOME.pq` | The target shape — one row per employee per rate type |
| `01_RAW_DATA.pq` | Source data in the paired column layout |
| `02_Solution1.pq` | Query for the paired layout (`RAW_DATA`) |
| `03_RAW_DATA2.pq` | Source data in the split column layout |
| `04_Solution2.pq` | Query for the split layout (`RAW_DATA2`) |

---

## How to use these files

### Step 1 — Open Power Query

Open a blank Excel workbook or a new Power BI Desktop file and open the Power Query Editor.

### Step 2 — Create a blank query for each file

For each `.pq` file you want to use:

1. In the Power Query Editor, go to **Home → New Source → Blank Query**
2. Open the **Advanced Editor**
3. Replace everything there with the contents of the `.pq` file
4. Name the query to match the file name without file order prefix (e.g. `RAW_DATA`, `RAW_DATA2`)

Load all queries you need before running the solution. The solution queries reference the source queries by name.

### Step 3 — Load in this order

Load the queries in this order so each one can find its dependencies:

1. `EXPECTED_OUTCOME` — standalone, no dependencies
2. `RAW_DATA` — standalone, no dependencies
3. `RAW_DATA2` — standalone, no dependencies
4. `Solution1` — references `RAW_DATA`
5. `Solution2` — references `RAW_DATA2`

---

## Guided walkthrough

Follow these steps alongside the video to build the solution from scratch.

### Understand the source layout

Open `RAW_DATA`. Notice the column structure:

- The first three columns identify the employee: `Employee`, `Company`, `Work date`
- The remaining columns are overtime data — three pairs of `OT Hours @x` and `OT Cost @x`
- Every pair shares the same rate suffix after the `@` symbol
- Hours always appears before Cost within each pair

This consistency is what makes a dynamic solution possible.

Open `EXPECTED_OUTCOME` and compare. The goal is one row per employee per rate type, with `Rate Type`, `Hours`, and `Cost` as separate columns.

---

### Build Solution — paired layout

Create a new blank query named `mySolution` and open the Advanced Editor. Follow along with the video.

**Step 1 — Identify the overtime columns**

```m
otCols = List.Select(Table.ColumnNames(Source), each Text.StartsWith(_, "OT "))
```

`Table.ColumnNames` returns all column names. `List.Select` filters that list, keeping only names that start with `"OT "`.

Inspect `otCols` — you should see a flat list of six column names in the order they appear in the table.

**Step 2 — Split the list into pairs**

```m
colSets = List.Split(otCols, 2)
```

`List.Split` divides the flat list into sublists of two items. With six column names and a chunk size of two, the result is three inner lists — one per rate type, each containing the Hours column name and the Cost column name for that rate.

Inspect `colSets` — the shape changes from a flat list to a list of lists; three list items that contain two text items each.

**Step 3 — Collect the values from each pair**

```m
pairs = Table.AddColumn(Source, "_rows", each
    List.Transform(colSets, (c) =>
        List.Transform(c, (v) => Record.Field(_, v))
    )
)
```

For every row in the table, the outer `List.Transform` iterates over `colSets`. For each pair of column names `c`, the inner `List.Transform` retrieves the value for each column name `v` from the current table row `_` using `Record.Field`.

Inspect `_rows` — each cell contains a list of inner lists. Drill into one cell to see the values.

**Step 4 — Add the rate type label**

```m
{"OT @" & Text.AfterDelimiter(List.First(c), "@")}
```

`List.First` returns the first column name in the pair. `Text.AfterDelimiter` extracts the text after the `@` symbol. The `"OT @"` prefix recreates the label in a consistent format.

This label is placed in front of the two collected values using `&` to form a single list of three items: label, Hours, Cost — always in that order.

**Step 5 — Turn each inner list into a record**

```m
Record.FromList(
    {"OT @" & Text.AfterDelimiter(List.First(c), "@")} &
    List.Transform(c, (v) => Record.Field(_, v)),
    {"Rate Type", "Hours", "Cost"}
)
```

`Record.FromList` pairs each value with a field name. The first argument is the list of values; the second is the list of field names. The field names are hard-coded here because every set produces exactly three values in a known order.

Inspect `_rows` again — each cell now contains a list of records instead of a list of lists.

**Step 6 — Ascribe the column type**

Add the optional type argument to `Table.AddColumn`:

```m
type {[Rate Type = text, Hours = number, Cost = number]}
```

This ascribes a type to the new column — it declares what the column contains without value conversion or type check. If the source values are already the correct types, this is fine. If Hours or Cost arrive as text from the source, you will need a conversion step. See [Adapting this for your data](#adapting-this-for-your-data) below.

**Step 7 — Remove the original OT columns, then expand**

```m
removeCols = Table.RemoveColumns(pairs, otCols),
expandLists = Table.ExpandListColumn(removeCols, "_rows"),
expandRecords = Table.ExpandRecordColumn(expandLists, "_rows", {"Rate Type", "Hours", "Cost"})
```

`Table.RemoveColumns` uses the same `otCols` list to remove all OT source columns.

`Table.ExpandListColumn` turns each list in `_rows` into separate rows — the row count multiplies by the number of rate types.

`Table.ExpandRecordColumn` turns the record fields into columns.

Compare the final result to `EXPECTED_OUTCOME`.

---

### Build Solution2 — split layout

Open `RAW_DATA2` and compare the column order to `RAW_DATA`. All Hours columns now appear before all Cost columns.

Duplicate the `mySolution` query and rename it `mySolution2`. Update the source reference to `RAW_DATA2`.

The only step that needs to change is `colSets`:

```m
colSets = List.Zip(List.Split(otCols, List.Count(otCols) / 2))
```

`List.Split(otCols, List.Count(otCols) / 2)` divides the flat list into two equal groups — the first half contains all Hours column names, the second contains all Cost column names.

`List.Zip` then combines the two lists, pairing the first item from each list, then the second, and so on.

The result is the same set of pairs as `Solution1` — the rest of the query is unchanged.

If you want to understand `List.Zip` in more detail, there is a dedicated video on the channel: [List.Zip Explained](https://www.youtube.com/watch?v=Ei5uDoyBNV0).

---

## Adapting this for your data


- **The column prefix is different**  
Replace `"OT "` in `Text.StartsWith(_, "OT ")` with the prefix your columns use. The space after the letters is intentional — adjust it to match your naming convention.

- **The delimiter between column name and rate label is different**  
Replace `"@"` in `Text.AfterDelimiter` with the delimiter your column names use.

- **The number of column pairs is different**  
No change needed. The solution reads the number of pairs from the data. Add or remove rate columns in the source and the solution adapts — as long as the naming convention and pairing order stay consistent.

- **The source query name is different**  
Replace `RAW_DATA` or `RAW_DATA2` in the `Source =` line with the name of your actual source query.

- **Hours or Cost arrive as text**  
Type ascription will not convert values. If your source delivers numeric columns as text, add a `Table.TransformColumnTypes` step after the expand steps to convert `Hours` and `Cost` to `number`.

- **You have more than two fields per group**  
The solution assumes each group produces exactly three output values: a label, Hours, and Cost. If your export has a different structure — for example three columns per rate — adjust the number of column names retrieved per inner list, the label extraction, and the field names passed to `Record.FromList`.

---

## Important note

These queries are for learning and practice. Review and test any query with your own data before using it in a business process.

---

## Resources
- Full Step-by-Step Guide: See sections above
- Channel: [YouTube](https://www.youtube.com/@melissa_de_korte)
- Questions? [Connect on LinkedIn](https://www.linkedin.com/in/melissa-de-korte-64585884)
