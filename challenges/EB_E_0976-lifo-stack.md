# Challenge 976: LIFO Stack

## Overview

This challenge is about simulating a Last-In-First-Out (LIFO) stack data structure using three operations.

Each row in the input table contains an operation:

- `PUSH` — add a value to the top of the stack
- `POP` — remove and discard the top element from the stack (if the stack is not empty)
- `MAX` — query the maximum value currently in the stack

The goal is to process each operation in order and return the maximum value for every `MAX` operation. For `PUSH` and `POP` operations, the answer is blank.

## Challenge rules

Simulate a LIFO stack with these requirements:

1. `PUSH` adds the value from the `Value` column to the top of the stack
2. `POP` removes the top element from the stack (only if the stack is not empty)
3. `MAX` returns the maximum value currently in the stack
4. For `PUSH` and `POP` operations, the `Answer Expected` column should be blank
5. For `MAX` operations, the `Answer Expected` column should contain the maximum value in the stack at that point (or blank if the stack is empty)
6. Operations are processed sequentially from top to bottom

## Source

Challenge creator: [Excel BI (Vijay Verma) on LinkedIn](https://www.linkedin.com/in/excelbi)

Challenge file: [Download practice file](https://lnkd.in/dgdEV5WT)

## Sample data

Input table:

| Operation | Value |
|---|---:|
| `PUSH` | `12` |
| `PUSH` | `5` |
| `MAX` |  |
| `PUSH` | `8` |
| `PUSH` | `19` |
| `MAX` |  |
| `POP` |  |
| `MAX` |  |
| `PUSH` | `3` |
| `PUSH` | `15` |
| `MAX` |  |
| `POP` |  |
| `POP` |  |
| `MAX` |  |
| `PUSH` | `22` |
| `PUSH` | `7` |
| `PUSH` | `14` |
| `MAX` |  |
| `POP` |  |
| `POP` |  |
| `MAX` |  |
| `POP` |  |
| `POP` |  |
| `POP` |  |
| `MAX` |  |

Expected result:

| Operation | Value | Answer Expected |
|---|---:|---:|
| `PUSH` | `12` |  |
| `PUSH` | `5` |  |
| `MAX` |  | `12` |
| `PUSH` | `8` |  |
| `PUSH` | `19` |  |
| `MAX` |  | `19` |
| `POP` |  |  |
| `MAX` |  | `12` |
| `PUSH` | `3` |  |
| `PUSH` | `15` |  |
| `MAX` |  | `15` |
| `POP` |  |  |
| `POP` |  |  |
| `MAX` |  | `12` |
| `PUSH` | `22` |  |
| `PUSH` | `7` |  |
| `PUSH` | `14` |  |
| `MAX` |  | `22` |
| `POP` |  |  |
| `POP` |  |  |
| `MAX` |  | `22` |
| `POP` |  |  |
| `POP` |  |  |
| `POP` |  |  |
| `MAX` |  | `12` |

## Power Query M solution

```m
let
    S = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    A = List.Accumulate(
        Table.ToRecords(S),
        [Stack = {}, Out = {}],
        (x, r) => [
            Stack = if r[Operation] = "PUSH" 
                then {r[Value]} & x[Stack] 
                else if r[Operation] = "POP" 
                then List.Skip(x[Stack]) 
                else x[Stack],
            Out = x[Out] & {if r[Operation] = "MAX" then List.Max(Stack, null) else null}
        ]
    ),
    Result = Table.FromColumns(
        Table.ToColumns(S) & {A[Out]},
        Table.ColumnNames(S) & {"Answer"}
    )
in
    Result
```

## Breakdown

The solution uses `List.Accumulate` to process each row sequentially, maintaining the stack state throughout the operations.

---

```m
S = Excel.CurrentWorkbook(){[Name="Table1"]}[Content]
```

This gets the source table named `Table1` from the current Excel workbook.

The table is stored in variable `S` for use in later steps.

---

```m
A = List.Accumulate(
    Table.ToRecords(S),
    [Stack = {}, Out = {}],
    (x, r) => ...
)
```

`List.Accumulate` processes each row of the table sequentially while maintaining state.

The accumulation uses three parts:

| Part | Purpose |
|---|---|
| `Table.ToRecords(S)` | converts the table into a list of records, one per row |
| `[Stack = {}, Out = {}]` | initial state: empty stack and empty output list |
| `(x, r) => ...` | for each row record `r`, update the state `x` |

The initial state contains:

- `Stack` — an empty list that represents the stack
- `Out` — an empty list that collects the answers for `MAX` operations

---

```m
Stack = if r[Operation] = "PUSH" 
    then {r[Value]} & x[Stack] 
    else if r[Operation] = "POP" 
    then List.Skip(x[Stack]) 
    else x[Stack]
```

This updates the stack based on the current operation.

For `PUSH`:

```m
{r[Value]} & x[Stack]
```

A new list is created with the `Value` from the current row at the front, followed by the existing stack.

For example, if the current stack is `{12, 5}` and the operation is `PUSH 8`, the result is:

```m
{8, 12, 5}
```

The `&` operator concatenates two lists, placing the new value at the beginning (top of the stack).

---

For `POP`:

```m
List.Skip(x[Stack])
```

`List.Skip` removes the first element from the stack (the top).

If the current stack is `{8, 12, 5}` and the operation is `POP`, the result is:

```m
{12, 5}
```

If the stack is already empty, `List.Skip` on an empty list returns an empty list, so no error occurs.

---

For `MAX`:

```m
else x[Stack]
```

The stack is unchanged. A `MAX` query does not modify the stack.

---

```m
Out = x[Out] & {if r[Operation] = "MAX" then List.Max(Stack, null) else null}
```

This appends a new item to the output list.

For each row:

- If the operation is `MAX`, the expression calculates `List.Max(Stack, null)`
- Otherwise, it appends `null`

`List.Max(Stack, null)` returns the maximum value in the current stack. The `null` parameter is the value to return if the stack is empty.

For example:

| Stack | Operation | Appended value |
|---|---|---|
| `{12, 5}` | `MAX` | `12` |
| `{19, 8, 12, 5}` | `MAX` | `19` |
| `{22, 7, 14}` | `MAX` | `22` |
| `{}` | `MAX` | `null` |

The appended value is wrapped in `{ }` to create a single-item list, which is then concatenated to `x[Out]`.

---

```m
Result = Table.FromColumns(
    Table.ToColumns(S) & {A[Out]},
    Table.ColumnNames(S) & {"Answer"}
)
```

This reconstructs the table with the new `Answer` column added.

`Table.ToColumns(S)` converts the source table into a list of column lists:

```m
{
    {value1, value2, value3, ...},  // Operation column
    {value1, value2, value3, ...}   // Value column
}
```

`A[Out]` is the list of answer values calculated by the accumulation:

```m
{null, null, 12, null, null, 19, null, 12, ...}
```

These are combined with the `&` operator:

```m
Table.ToColumns(S) & {A[Out]}
```

This produces a list of all columns (original plus the new `Answer` column).

---

```m
Table.ColumnNames(S) & {"Answer"}
```

This creates the column name list for the result table.

`Table.ColumnNames(S)` returns the names of the original columns:

```m
{"Operation", "Value", "Answer Expected"}
```

A new name `"Answer"` is added to the end:

```m
{"Operation", "Value", "Answer Expected", "Answer"}
```

---

```m
in
    Result
```

The query returns the result table with the calculated `Answer` column.

## Detailed example

To understand how the stack evolves, walk through the first few operations:

**Initial state:**
- Stack: `{}`
- Out: `{}`

**Row 1: PUSH 12**
- Operation is `PUSH`, so add `12` to the stack
- New stack: `{12}`
- Output: append `null`
- Out: `{null}`

**Row 2: PUSH 5**
- Operation is `PUSH`, so add `5` to the top
- New stack: `{5, 12}` (5 is at the top, 12 below)
- Output: append `null`
- Out: `{null, null}`

**Row 3: MAX**
- Operation is `MAX`, so query the maximum
- Stack remains: `{5, 12}`
- `List.Max({5, 12})` returns `12`
- Output: append `12`
- Out: `{null, null, 12}`

**Row 4: PUSH 8**
- Operation is `PUSH`, so add `8` to the top
- New stack: `{8, 5, 12}`
- Output: append `null`
- Out: `{null, null, 12, null}`

**Row 5: PUSH 19**
- Operation is `PUSH`, so add `19` to the top
- New stack: `{19, 8, 5, 12}`
- Output: append `null`
- Out: `{null, null, 12, null, null}`

**Row 6: MAX**
- Operation is `MAX`, so query the maximum
- Stack remains: `{19, 8, 5, 12}`
- `List.Max({19, 8, 5, 12})` returns `19`
- Output: append `19`
- Out: `{null, null, 12, null, null, 19}`

**Row 7: POP**
- Operation is `POP`, so remove the top element
- Stack becomes: `{8, 5, 12}` (19 is removed)
- Output: append `null`
- Out: `{null, null, 12, null, null, 19, null}`

**Row 8: MAX**
- Operation is `MAX`, so query the maximum
- Stack remains: `{8, 5, 12}`
- `List.Max({8, 5, 12})` returns `12`
- Output: append `12`
- Out: `{null, null, 12, null, null, 19, null, 12}`

This pattern continues through all remaining rows.

## Why this works

The key insight is that `List.Accumulate` maintains running state across multiple iterations.

For each row:

1. The current stack state is available in `x[Stack]`
2. A modified stack is calculated based on the operation
3. The output value (or null) is calculated and appended
4. The new state becomes `[Stack = ..., Out = ...]` for the next iteration

This allows the solution to:

- Maintain a list representing the stack
- Add elements to the front (PUSH)
- Remove elements from the front (POP)
- Query the maximum without modifying the stack (MAX)
- Track all answers in the `Out` list

At the end, the `Out` list contains one entry per row, with answers for `MAX` rows and nulls for `PUSH`/`POP` rows.

The reconstructed table combines the original columns with the new `Answer` column, producing the final result.

## Notes

This solution assumes:

- the Excel table is named `Table1`
- the operation column is named `Operation`
- the value column is named `Value`
- the stack is represented as a list with the top element at position 0

If your table has a different name, update this part:

```m
Excel.CurrentWorkbook(){[Name="Table1"]}[Content]
```

If your column names are different, update the field references in the operation checks:

```m
r[Operation]
r[Value]
```

The solution creates a column named `Answer`. If needed, this can be renamed to match the workbook's expected output column.

## Tags

`Power Query` · `M code` · `List.Accumulate` · `Data structures` · `Stack simulation` · `List.Max` · `Table.FromColumns` · `Excel`
