# Exploring Table.Pivot in Power Query M

`Table.Pivot` reshapes tall, narrow data — one row per combination — into a compact grid where patterns are easier to spot. Any dataset that repeats a key for every associated item is a candidate.

This video walks through the Pivot Column dialog, maps UI options to the M code it generates, and builds six different outputs from the same source data.

Video: [Watch on YouTube](https://youtu.be/aGRG59l_MQA)

LinkedIn: [Connect with Melissa](https://www.linkedin.com/in/melissa-de-korte/)


## Source data

The sample data is an API access list.

```m
// RAW_DATA
let
    Source = Table.FromRows(
        {
            {"Wendy", "Graph", "Admin"},
            {"Wendy", "GitHub", "Admin"},
            {"Wendy", "Salesforce", "Admin"},
            {"Bell", "Jira", "Read/Write"},
            {"Lily", "Salesforce", "Read"},
            {"Jane", "Graph", "Read"},
            {"Jane", "GitHub", "Read"},
            {"Moira", "", ""}
        },
        type table [User = text, API = text, Role = text]
    )[[User], [API]]
in
    Source
```

The sample looks like this.  
Two columns: each row represents one user–API assignment.

| User | API |
|---|---|
| Wendy | Graph |
| Wendy | GitHub |
| Wendy | Salesforce |
| Bell | Jira |
| Lily | Salesforce |
| Jane | Graph |
| Jane | GitHub |
| Moira | |

---

## Data quality

Three things to check before pivoting:

**Nulls.** Moira's API column contains an empty string, not a null. Nulls in the pivot column cause problems. Replace them with something meaningful or filter out rows that aren't relevant.

**Normalize text.** Leading or trailing spaces, inconsistent casing — Power Query treats every variation as a separate value. Anything that should be considered equal must match exactly.

**Duplicate rows.** Decide whether you need distinct combinations only. Duplicates affect counts and can trigger errors when no aggregation is specified.

## From UI to M code

The Pivot Column dialog maps directly to the parameters of `Table.Pivot`.

<img width="485" height="649" alt="PivotDialog" src="https://github.com/user-attachments/assets/68c7b4fc-45f7-4cc4-b3e3-09b43a36e625" />


### How the dialog maps to the syntax

```m
Table.Pivot( table, pivotValues, attributeColumn, valueColumn, aggregationFunction )
```

| # | UI element | Parameter | What it controls |
|---|---|---|---|
| **①** | Selected column (API) | `pivotValues` + `attributeColumn` | The column you select before invoking the Pivot. Its distinct values become the new column headers (`pivotValues`), and it is named as the `attributeColumn`. |
| **②③** | Transform tab → Pivot Column button | — | The route to invoke the dialog. |
| **④** | Values Column dropdown (User) | `valueColumn` | The column whose values fill the cells in the grid. |
| **⑤** | Advanced options → Aggregate Value Function | `aggregationFunction` | Controls how multiple values at the same intersection are handled. Selecting "Don't Aggregate" omits the fifth parameter entirely. |

Any column not consumed by `attributeColumn` or `valueColumn` becomes the row axis — it groups the rows automatically.

## The six outputs

All six outputs are produced from the same source data. Each demonstrates a different combination of parameters.

---

### Output 1 — Combined text (Text.Combine)

No leftover columns, so the result is a single row. The aggregation function joins all values at each intersection with a comma and space.

```m
Table.Pivot(
    Source,
    List.Distinct(Source[API]),
    "API", "User",
    each Text.Combine(_, ", ")
)
```

<img width="843" height="57" alt="output-text-combine" src="https://github.com/user-attachments/assets/b857fdb3-5f17-4f62-9148-612314b6d120" />


---

### Output 2 — Expanded rows from a single-row pivot

When a pivot produces exactly one row, the lists in each cell can be expanded back into rows.

```m
Pivot = Table.Pivot(Source, List.Distinct(Source[API]), "API", "User", each _),
Expand = Table.FromColumns(Record.ToList(Pivot{0}), Table.ColumnNames(Pivot))
```

<img width="851" height="78" alt="output-expanded-rows" src="https://github.com/user-attachments/assets/a8ac5086-87d7-43ee-af5b-808be68ea68d" />


This technique depends on the pivot producing exactly one row. If there are multiple rows, all except the first are discarded.

---

### Output 3 — Same-column pattern (User values in the grid)

Setting API as both `attributeColumn` and `valueColumn` frees up the User column for the row axis. When User is the first column of the source table, user names appear in the grid.

```m
Table.Pivot(Source, List.Distinct(Source[API]), "API", "API")
```

<img width="1001" height="135" alt="output-same-col-user-values" src="https://github.com/user-attachments/assets/ebc4e576-30d1-45e1-9edc-0e4305fe0e51" />


---

### Output 4 — Same-column pattern (API values in the grid)

The exact same expression, but this time the source table has API as its first column. Now API values appear in the grid instead.

```m
Table.Pivot(Source, List.Distinct(Source[API]), "API", "API")
```

<img width="994" height="136" alt="output-same-col-api-values" src="https://github.com/user-attachments/assets/23a155fb-a4b2-4806-bd46-99c063b62253" />


**The first-column rule:** when `attributeColumn` and `valueColumn` reference the same column, Power Query uses the first column of the input table to determine what fills the cells. This behavior appears to be undocumented.

---

### Output 5 — Boolean grid using a custom column

A helper column set to `true` is added before pivoting. The grid shows TRUE where access exists and null elsewhere.

```m
eachTRUE = Table.AddColumn(Source, "Custom", each true, type logical),
Pivot = Table.Pivot(eachTRUE, List.Distinct(eachTRUE[API]), "API", "Custom")
```

<img width="999" height="141" alt="output-true-null" src="https://github.com/user-attachments/assets/630dd9be-4288-46f1-8789-8254af4f4140" />


---

### Output 6 — Boolean grid using an aggregation function

No extra columns needed. `List.NonNullCount` produces 0/1 values, and `Logical.From` converts them to TRUE/FALSE.

```m
Table.Pivot(
    Source,
    List.Distinct(Source[API]),
    "API", "API",
    each Logical.From(List.NonNullCount(_))
)
```

<img width="997" height="137" alt="output-true-false" src="https://github.com/user-attachments/assets/70330eb6-711f-4743-927e-32a628efd6b7" />


---

## Key concepts

### The error without aggregation

When no `aggregationFunction` is provided (the dialog's "Don't Aggregate" option), `Table.Pivot` expects exactly one value at every intersection. If more than one value exists, it raises the error: *"There were too many elements in the enumeration to complete the operation."*

### What the aggregation function receives

The aggregation function always receives a **list** — the collected values for that intersection. Any function that processes a list and returns a scalar value can be used: `List.Sum`, `List.Count`, `List.NonNullCount`, `Text.Combine`, or a custom function.

Passing `each _` returns the list unchanged, which is useful for inspecting what's collected before choosing an aggregation.

### Leftover columns create the row axis

Any column not named as `attributeColumn` or `valueColumn` remains in the output and groups the rows. This happens automatically — there is no explicit "group by" parameter.

## What to take away

The UI gets you started. The formula bar lets you go further. The same-column pattern, custom aggregations, boolean grids — all of that comes from understanding what the five parameters do.

## Tags

`Power Query` · `M code` · `Table.Pivot` · `List.Distinct` · `Text.Combine` · `List.NonNullCount` · `Logical.From` · `Table.AddColumn` · `Table.FromColumns` · `Excel` · `Power BI`
