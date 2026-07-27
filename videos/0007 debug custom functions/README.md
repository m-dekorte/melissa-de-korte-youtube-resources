# Debugging Custom Functions in Power Query M

A step-by-step guide, these techniques work with any custom function in Power Query M, whether you're using Excel or Power BI.

- **Video:** [Watch on YouTube](https://youtu.be/wkIo9fOrSeE)
- **LinkedIn:** [Connect with Melissa](https://www.linkedin.com/in/melissa-de-korte-64585884/)

- **Sample:** [videos/0006 multi row headers/01_RAW_DATA.pq](https://github.com/m-dekorte/melissa-de-korte-youtube-resources/blob/main/videos/0006%20multi%20row%20headers/01_RAW_DATA.pq)

---

## Overview

Custom functions in Power Query are powerful, but they can be difficult to debug. The Preview Pane displays parameters and an option to invoke it. But when something goes wrong, you need a way to see what is happening inside.

This guide covers practical techniques for inspecting a custom function's internals. Techniques that can be combined as needed.

> **Tip!**  
> Make a copy of your function query before editing. This gives you a backup you can use to restore earlier code sections by copying and pasting them back.  Create new backups throughout development whenever you reach an important milestone.  

The running example throughout this guide is a function that promotes multiline headers:

```powerquery
(t as table, startRow as number, rowCount as number) as table =>
let
    offset = startRow-1,
    headers = List.Transform(Table.ToColumns(Table.Range(t, offset, rowCount)), each
        let x = List.RemoveMatchingItems(_, {"", null}) in
        if List.IsEmpty(x) then null else Text.Combine(List.Transform(x, Text.From), " ")
    ),
    transform = Table.PromoteHeaders(
        Table.InsertRows(Table.Skip(t, offset+rowCount), 0,
            {Record.FromList(headers, Table.ColumnNames(t))}
        )
    )
in
    transform
```

This function has three parameters (`t`, `startRow`, `rowCount`), a return type (`as table`), and several nested expressions.

---

## Disable the function

A function cannot be stepped through like a regular query. Power Query evaluates the query as a function value, the Preview pane displays its parameters and an option to invoke it. Once invoked, the calling query shows only the value returned by the function, not the intermediate steps evaluated to obtain that result.

Therefore to inspect the internals, the first step is to temporarily disable the function by commenting out the function declaration:

```powerquery
/* (t as table, startRow as number, rowCount as number) as table => */
let
    offset = startRow-1,
    headers = List.Transform(Table.ToColumns(Table.Range(t, offset, rowCount)), each
        let x = List.RemoveMatchingItems(_, {"", null}) in
        if List.IsEmpty(x) then null else Text.Combine(List.Transform(x, Text.From), " ")
    ),
    transform = Table.PromoteHeaders(
        Table.InsertRows(Table.Skip(t, offset+rowCount), 0,
            {Record.FromList(headers, Table.ColumnNames(t))}
        )
    )
in
    transform
```

The code inside the `let` expression is unchanged. Only the function has been disabled.

At this point, the query will error because `t`, `startRow`, and `rowCount` are no longer defined. They were parameters, now they are references to unknown names. The next step fixes that.

---

## Parameter values

With the function disabled, the parameter names have no values. Replace them by adding variables immediately below the `let` keyword:

```powerquery
/* (t as table, startRow as number, rowCount as number) as table => */
let
    /* PARAM */
    t = RAW_DATA,
    startRow = 6,
    rowCount = 3,

    /* FN BODY */
    offset = startRow-1,
    headers = List.Transform(Table.ToColumns(Table.Range(t, offset, rowCount)), each
        let x = List.RemoveMatchingItems(_, {"", null}) in
        if List.IsEmpty(x) then null else Text.Combine(List.Transform(x, Text.From), " ")
    ),
    transform = Table.PromoteHeaders(
        Table.InsertRows(Table.Skip(t, offset+rowCount), 0,
            {Record.FromList(headers, Table.ColumnNames(t))}
        )
    )
in
    transform
```

Replace `RAW_DATA` with a reference to the actual table you want to test with. This could be a query name, or a table expression. For example:

```powerquery
t = Excel.CurrentWorkbook(){[Name = "Table1"]}[Content],
```

Now the query runs as a regular `let` expression with fixed inputs. Power Query evaluates every step and the Applied Steps pane shows each variable. You can click through every step including `offset`, `headers`, and `transform` to inspect them one by one.

This is the simplest way to turn a function into a debuggable query.

---

## Using a record expression

A `let` expression only shows the value named after `in` or the step that is selected in the Applied Steps pane. To review everything at the same time, convert it to a record expression.

A record expression uses square brackets instead of `let ... in` and returns all fields, the full record:

```powerquery
/* (t as table, startRow as number, rowCount as number) as table => */
[
    /* PARAM */
    t = RAW_DATA,
    startRow = 6,
    rowCount = 3,

    /* FN BODY */
    offset = startRow-1,
    headers = List.Transform(Table.ToColumns(Table.Range(t, offset, rowCount)), each
        let x = List.RemoveMatchingItems(_, {"", null}) in
        if List.IsEmpty(x) then null else Text.Combine(List.Transform(x, Text.From), " ")
    ),
    transform = Table.PromoteHeaders(
        Table.InsertRows(Table.Skip(t, offset+rowCount), 0,
            {Record.FromList(headers, Table.ColumnNames(t))}
        )
    )
]
```

The Preview Pane now shows a record that displays every field. Click any field value to drill down into it or click the whitespace to see a secondary preview at the bottom of the pane:

| Field | Type you will see |
|---|---|
| `t` | Table |
| `startRow` | Number |
| `rowCount` | Number |
| `offset` | Number |
| `headers` | List |
| `transform` | Table |

This is especially useful when you want to see the scope of an error.

A record and a `let` expression work similarly in that each field can reference other fields. The practical difference for debugging is that a record exposes everything, while `let ... in` only exposes one value at a time.

---

## Break apart nested expressions

Nested expressions are compact but harder to debug because you cannot see intermediate values. The solution is to extract each step and assign it to its own variable, working from the inside out.

### A nested expression

The `headers` step contains three levels of nesting:

```powerquery/
headers = List.Transform(Table.ToColumns(Table.Range(t, offset, rowCount)), each
    let x = List.RemoveMatchingItems(_, {"", null}) in
    if List.IsEmpty(x) 
    then null 
    else Text.Combine(List.Transform(x, Text.From), " ")
)
```

Reading from the inside out, this does:

1. `Table.Range(t, offset, rowCount)` — extract the header rows
2. `Table.ToColumns(...)` — transpose rows into column-based lists
3. `List.Transform(...)` — process each column's list into a header name

Each step feeds directly into the next. If something goes wrong, it can be hard to tell what caused it.

### An unnested expression

Extract each separate expression and assign it to its own variable:

```powerquery
/* (t as table, startRow as number, rowCount as number) as table => */
[
    /* PARAMS */
    t = RAW_DATA,
    startRow = 6,
    rowCount = 3,

    /* FN BODY */
    offset = startRow-1,
    headerRows = Table.Range(t, offset, rowCount),
    headerToCols = Table.ToColumns(headerRows),
    headers = List.Transform(headerToCols, each
        let x = List.RemoveMatchingItems(_, {"", null}) in
        if List.IsEmpty(x) 
        then null 
        else Text.Combine(List.Transform(x, Text.From), " ")
    ),
    transform = Table.PromoteHeaders(
        Table.InsertRows(Table.Skip(t, offset+rowCount), 0,
            {Record.FromList(headers, Table.ColumnNames(t))}
        )
    )
]
```

Now you can inspect each layer separately:

| Field | What you see |
|---|---|
| `headerRows` | A table with only the header rows extracted |
| `headerToCols` | A list of lists, one per column |
| `headers` | A list of combined header names (text or null) |


### Going deeper

If the `List.Transform` step itself needs debugging, isolate what happens to one column by pulling a single item from the list and processing it:

```powerquery
/* (t as table, startRow as number, rowCount as number) as table => */
[
    /* PARAM */
    t = RAW_DATA,
    startRow = 6,
    rowCount = 3,

    /* FN BODY */
    offset = startRow-1,
    headerRows = Table.Range(t, offset, rowCount),
    headerToCols = Table.ToColumns(headerRows),
    headers = List.Transform(headerToCols, each
        let x = List.RemoveMatchingItems(_, {"", null}) in
        if List.IsEmpty(x) 
        then null 
        else Text.Combine(List.Transform(x, Text.From), " ")
    ),

    /* Change the index to test a different column */
    singleColumn = headerToCols{7},
    cleaned = List.RemoveMatchingItems(singleColumn, {"", null}),
    toText = List.Transform(cleaned, Text.From),
    combined = Text.Combine(toText, " "),

    transform = Table.PromoteHeaders(
        Table.InsertRows(Table.Skip(t, offset+rowCount), 0,
            {Record.FromList(headers, Table.ColumnNames(t))}
        )
    )
]
```

This adds four fields that trace one column through every stage:

| Field | Value for column index `7` |
|---|---|
| `singleColumn` | `{"Net", "sales", "amount"}` |
| `cleaned` | `{"Net", "sales", "amount"}` |
| `toText` | `{"Net", "sales", "amount"}` |
| `combined` | `Net sales amount` |

Change the index to test other columns. For example, `headerToCols{0}` returns a column where all values are null, which follows the `null` branch of the condition.

### Unnesting the transform step

The same approach applies to the `transform` step. It nests three function calls:

```powerquery
transform = Table.PromoteHeaders(
    Table.InsertRows(Table.Skip(t, offset+rowCount), 0,
        {Record.FromList(headers, Table.ColumnNames(t))}
    )
)
```

Unnested:

```powerquery
tblBody = Table.Skip(t, offset+rowCount),
headerRow = Record.FromList(headers, Table.ColumnNames(t)),
withHeaderRow = Table.InsertRows(tblBody, 0, {headerRow}),
transform = Table.PromoteHeaders(withHeaderRow)
```

Now you can verify each stage:

| Field | What to check |
|---|---|
| `tblBody` | Does it contain only data rows from the table body? |
| `headerRow` | Do record fields show expected results? |
| `withHeaderRow` | Are headers inserted at the top? |
| `transform` | Does the table match the expected outcome? |

### Final version

Every intermediate value is now visible in the Preview Pane. Once you have found and fixed the issue, you can convert the query back into a function: 
1. Use field access `[transform]` to return the value of that field.  
2. Comment out the hard-coded parameters. 
3. Restore the function signature.

Combining all of the above, the fully unnested and restored function looks like this:

```powerquery
(t as table, startRow as number, rowCount as number) as table =>
[
    /* PARAM */
    /* t = RAW_DATA,
    startRow = 6,
    rowCount = 3, */

    /* FN BODY */
    // extract and transform header row
    offset = startRow-1,
    headerRows = Table.Range(t, offset, rowCount),
    headerToCols = Table.ToColumns(headerRows),
    headers = List.Transform(headerToCols, each
        let x = List.RemoveMatchingItems(_, {"", null}) in
        if List.IsEmpty(x) 
        then null 
        else Text.Combine(List.Transform(x, Text.From), " ")
    ),

     // inspect a single column (change index as needed)
    singleColumn = headerToCols{7},
    cleaned = List.RemoveMatchingItems(singleColumn, {"", null}),
    toText = List.Transform(cleaned, Text.From),
    combined = Text.Combine(toText, " "),


    // collect and promote
    tblBody = Table.Skip(t, offset+rowCount),
    headerRow = Record.FromList(headers, Table.ColumnNames(t)),
    withHeaderRow = Table.InsertRows(tblBody, 0, {headerRow}),
    transform = Table.PromoteHeaders(withHeaderRow)
][transform]
```

Like `let` expressions, `record` expressions are evaluated lazily. A field is evaluated only when its value is needed, so fields that do not contribute to the requested result may never be evaluated.

---

## Additional | How to inspect a specific element

When a step returns a list, you can access an individual item by appending an index within curly brackets `{ }`.

In the record expression, add a field that accesses a specific position:

```powerquery
thirdColumn = headerToCols{2}
```

Power Query (M) uses a zero-based index. `{0}` is the first item, `{2}` is the third.

For tables, use the same syntax to access a specific row:

```powerquery
firstDataRow = tblBody{0}
```

This returns a record representing that row.

---

---

## Summary

| Technique | What it does | When to use it |
|---|---|---|
| Disable the function | Comments out the function signature so the body runs as a regular query | Always — this is the starting point for any function debugging |
| Hardcode parameters | Assigns fixed values to parameter names | Always — needed after disabling the function |
| Record expression | Returns all fields as a navigable record | When you want to see multiple steps at once |
| Break apart nested expressions | Extract each layer of nesting into its own named step | When a nested expression errors or produces unexpected results |
| Inspect an item | Accesses a single item by position | When a list or table step works for most items but fails for some |

## Tags

`Power Query` · `M code` · `Debugging` · `Custom function` · `Record expression` · `Advanced Editor` · `Excel` · `Power BI`
