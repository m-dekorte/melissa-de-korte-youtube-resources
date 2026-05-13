# Challenge 401: Column Splitting

## Overview

This challenge is about splitting each value in an `ID` column into three separate columns:

- `ID1`
- `ID2`
- `ID3`

The goal is to split the text as evenly as possible from left to right.

## Challenge rules

Split each `ID` into three columns using these rules:

1. If the number of characters is divisible by 3, split the value into three equal parts.
2. Otherwise, distribute the characters as evenly as possible from left to right.
3. Extra characters should go to `ID1` first, then `ID2`.
4. The difference in length between the parts should be at most 1.

## Source

Challenge creator: [Omid Motamedisedeh on LinkedIn](https://www.linkedin.com/in/omidmot/)

Challenge file: [Download Excel file](https://lnkd.in/gvGf5ZpS)

## Sample data

Input table:

| ID |
|---|
| `SG2` |
| `FG1XY2` |
| `SN` |
| `MN11` |
| `FFGHI` |

Expected result:

| ID1 | ID2 | ID3 |
|---|---|---|
| `S` | `G` | `2` |
| `FG` | `1X` | `Y2` |
| `S` | `N` |  |
| `MN` | `1` | `1` |
| `FF` | `GH` | `I` |


## Power Query M solution

```powerquery
let
    Source = Excel.CurrentWorkbook(){[Name = "Table1"]}[Content],
    SplitID = Table.SplitColumn(
        Source,
        "ID",
        each 
            let
                n = Text.Length(_)
            in
                List.Transform(
                    {0..2},
                    (i) =>
                        Text.Range(
                            _,
                            Number.RoundUp(i * n / 3),
                            Number.RoundUp((i + 1) * n / 3) - Number.RoundUp(i * n / 3)
                        )
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

This splits the `ID` column. Instead of using a fixed delimiter or fixed number of characters, the split logic is defined by a custom function.

---

```powerquery
each 
    let
        n = Text.Length(_)
    in
        ...
```

For each value in the `ID` column, `Text.Length(_)` counts the number of characters.

The underscore `_` represents the current `ID` value being processed.

For example:

| ID | Length |
|---|---:|
| `SG2` | 3 |
| `FG1XY2` | 6 |
| `SN` | 2 |
| `MN11` | 4 |
| `FFGHI` | 5 |

---

```powerquery
List.Transform(
    {0..2},
    (i) => ...
)
```

This creates a list with three items by looping through the values `0`, `1`, and `2`.

Each value of `i` represents one of the output columns:

| `i` | Output column |
|---:|---|
| `0` | `ID1` |
| `1` | `ID2` |
| `2` | `ID3` |

---

```powerquery
Number.RoundUp(i * n / 3)
```

This calculates the starting position for each part.

The text is divided into three sections by calculating boundary positions across the length of the text. `Number.RoundUp` makes sure any extra characters are assigned from left to right.

For an ID with 5 characters, such as `FFGHI`, the boundaries are:

| Boundary | Calculation | Rounded up |
|---:|---:|---:|
| Start of `ID1` | `0 * 5 / 3` | `0` |
| Start of `ID2` | `1 * 5 / 3` | `2` |
| Start of `ID3` | `2 * 5 / 3` | `4` |
| End | `3 * 5 / 3` | `5` |

That gives these ranges:

| Part | Position range | Length | Result |
|---|---|---:|---|
| `ID1` | 0 to 2 | 2 | `FF` |
| `ID2` | 2 to 4 | 2 | `GH` |
| `ID3` | 4 to 5 | 1 | `I` |

So `FFGHI` becomes:

| ID1 | ID2 | ID3 |
|---|---|---|
| `FF` | `GH` | `I` |

---

```powerquery
Text.Range(
    _,
    Number.RoundUp(i * n / 3),
    Number.RoundUp((i + 1) * n / 3) - Number.RoundUp(i * n / 3)
)
```

`Text.Range` extracts each part of the text.

It uses three arguments:

1. The original text value.
2. The starting position.
3. The number of characters to return.

The number of characters is calculated by subtracting the current boundary from the next boundary.

---

```powerquery
{"ID1", "ID2", "ID3"}
```

This names the three new columns created by the split.

## Why this works

The key idea is to calculate the split positions instead of writing separate logic for each possible text length.

Using rounded-up boundary positions creates an even split while assigning extra characters from left to right.

Examples:

| Length | Split lengths |
|---:|---|
| 2 | 1, 1, 0 |
| 3 | 1, 1, 1 |
| 4 | 2, 1, 1 |
| 5 | 2, 2, 1 |
| 6 | 2, 2, 2 |
| 7 | 3, 2, 2 |
| 8 | 3, 3, 2 |

This matches the challenge rules because:

- equal splits happen automatically when the length is divisible by 3
- extra characters go to `ID1` first, then `ID2`
- the difference between the longest and shortest parts is never more than 1

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

`Power Query` · `M code` · `Text functions` · `Table.SplitColumn` · `Text.Range` · `Excel`
