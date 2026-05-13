# Challenge 411: Column Splitting

## Overview

This challenge is about splitting each value in an `ID` column into three separate columns:

- `ID1`
- `ID2`
- `ID3`

Instead of splitting the text into consecutive blocks, this challenge splits the text based on character positions.

## Challenge rules

Split each `ID` into three columns based on character positions:

1. `ID1` gets the characters in positions `1, 4, 7, 10, ...`
2. `ID2` gets the characters in positions `2, 5, 8, 11, ...`
3. `ID3` gets the characters in positions `3, 6, 9, 12, ...`
4. Positions are counted from left to right, starting at `1`.

## Source

Challenge creator: [Omid Motamedisedeh on LinkedIn](https://www.linkedin.com/in/omidmot/)

Challenge file: [Download Excel file](https://lnkd.in/gNUBTUh9)

## Sample data

Input table:

| ID |
|---|
| `QWERTYUIO` |
| `FG1XY2` |
| `SG2` |
| `SN` |
| `MN11` |
| `FFGHI` |

Expected result:

| ID1 | ID2 | ID3 |
|---|---|---|
| `QRU` | `WTI` | `EYO` |
| `FX` | `GY` | `12` |
| `S` | `G` | `2` |
| `S` | `N` |  |
| `M1` | `N` | `1` |
| `FH` | `FI` | `G` |

## Power Query M solution

```powerquery
let
    Source = Excel.CurrentWorkbook(){[Name = "Table1"]}[Content],
    SplitID = Table.SplitColumn(
        Source,
        "ID",
        each 
            List.Transform(
                List.Zip(
                    List.Split(
                        Text.ToList(_),
                        3
                    )
                ),
                Text.Combine
            ),
        {"ID1", "ID2", "ID3"}
    )
in
    SplitID
```

## Breakdown

The solution uses `Table.SplitColumn` with a custom splitting function.

```powerquery
Source = Excel.CurrentWorkbook(){[Name = "Table1"]}[Content]
```

This gets the source table named `Table1` from the current Excel workbook.

---

```powerquery
Table.SplitColumn(
    Source,
    "ID",
    each ...
)
```

This splits the `ID` column. Instead of splitting by delimiter or by a fixed number of characters, the split logic is created with a custom function.

---

```powerquery
Text.ToList(_)
```

`Text.ToList` converts the current `ID` value into a list of individual characters.

For example:

| ID | Resulting list |
|---|---|
| `FG1XY2` | `{ "F", "G", "1", "X", "Y", "2" }` |

The underscore `_` represents the current `ID` value being processed.

---

```powerquery
List.Split(
    Text.ToList(_),
    3
)
```

`List.Split` splits the list of characters into groups of three.

For example, `FG1XY2` becomes:

```powerquery
{
    {"F", "G", "1"},
    {"X", "Y", "2"}
}
```

Each inner list contains characters from one consecutive group of three.

---

```powerquery
List.Zip(...)
```

`List.Zip` combines items that are in the same position inside each inner list.

Using the previous example:

```powerquery
{
    {"F", "G", "1"},
    {"X", "Y", "2"}
}
```

becomes:

```powerquery
{
    {"F", "X"},
    {"G", "Y"},
    {"1", "2"}
}
```

This is the key step.

It changes the structure from consecutive groups into position-based groups:

| Output column | Characters collected |
|---|---|
| `ID1` | positions 1, 4, 7, 10, ... |
| `ID2` | positions 2, 5, 8, 11, ... |
| `ID3` | positions 3, 6, 9, 12, ... |

---

```powerquery
List.Transform(
    List.Zip(...),
    Text.Combine
)
```

`List.Transform` applies `Text.Combine` to each zipped list.

For `FG1XY2`, the zipped lists are:

```powerquery
{
    {"F", "X"},
    {"G", "Y"},
    {"1", "2"}
}
```

After `Text.Combine`, they become:

```powerquery
{
    "FX",
    "GY",
    "12"
}
```

These three values become the values for `ID1`, `ID2`, and `ID3`.

---

```powerquery
{"ID1", "ID2", "ID3"}
```

This names the three new columns created by the split.

---

## Detailed example

For this value:

```text
QWERTYUIO
```

The characters are first split into groups of three:

| Group | Characters |
|---:|---|
| 1 | `Q`, `W`, `E` |
| 2 | `R`, `T`, `Y` |
| 3 | `U`, `I`, `O` |

Then the groups are zipped by position:

| Output column | Characters | Result |
|---|---|---|
| `ID1` | `Q`, `R`, `U` | `QRU` |
| `ID2` | `W`, `T`, `I` | `WTI` |
| `ID3` | `E`, `Y`, `O` | `EYO` |

So `QWERTYUIO` becomes:

| ID1 | ID2 | ID3 |
|---|---|---|
| `QRU` | `WTI` | `EYO` |

## Why this works

The challenge is based on repeating character positions:

| Position | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 |
|---|---|---|---|---|---|---|---|---|---|
| Character | Q | W | E | R | T | Y | U | I | O |
| Output | ID1 | ID2 | ID3 | ID1 | ID2 | ID3 | ID1 | ID2 | ID3 |

The solution works because it first groups the text into chunks of three characters, then uses `List.Zip` to transpose those chunks.

This turns rows of three characters into three position-based columns.

In other words:

```text
Q W E
R T Y
U I O
```

becomes:

```text
Q R U
W T I
E Y O
```

That matches the required output.

## Notes

This solution assumes:

- the Excel table is named `Table1`
- the column to split is named `ID`

If your table has a different name, update this part:

```powerquery
Excel.CurrentWorkbook(){[Name = "Table1"]}[Content]
```

If your column has a different name, update this part:

```powerquery
"ID"
```

## Tags

`Power Query` · `M code` · `Text functions` · `Table.SplitColumn` · `Text.ToList` · `List.Split` · `List.Zip` · `Text.Combine` · `Excel`
