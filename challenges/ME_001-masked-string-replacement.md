# Masked String Replacement

## Overview

This challenge is about replacing asterisks in a masked text value with the corresponding characters from another column.

Each row contains:

- a masked string in `Masked String`
- replacement characters in `Chars`
- the expected completed text

The goal is to replace each `*` from left to right using the characters in `Chars`.

## Challenge rules

Replace the asterisks in `Masked String` with the corresponding characters from `Chars`.

For example:

| Masked String | Chars | Expected Answer |
|---|---|---|
| `*ave *u*!` | `HFn` | `Have Fun!` |

The first `*` is replaced by `H`, the second by `F`, and the third by `n`.

## Source

Challenge creator: [Excel BI (Vijay Verma) on LinkedIn](https://www.linkedin.com/in/excelbi/)

Challenge file: [Download workbook](https://1drv.ms/x/s!Akd5y6ruJhvhv0YU5Pi_XwvFOXSK)

Rerun by: [Meganathan Elumalai on LinkedIn](https://www.linkedin.com/in/meganathan-elumalai-a3449221/)

File: [Download workbook](https://lnkd.in/grTcP5g6)

## Sample data

| Masked String | Chars | Expected Answer |
|---|---|---|
| `To Kill a *ockin*bird` | `Mg` | `To Kill a Mockingbird` |
| `The G*eat G***by` | `rats` | `The Great Gatsby` |
| `*lys***` | `Uses` | `Ulysses` |
| `The Ca*cher in *** Rye` | `tthe` | `The Catcher in the Rye` |
| `*ride *nd *rejudic*` | `PaPe` | `Pride and Prejudice` |
| `*dv*ntures of Huckleberry ****` | `AeFinn` | `Adventures of Huckleberry Finn` |
| `Alice’s A***nture in ******land` | `dveWonder` | `Alice’s Adventure in Wonderland` |
| `** the Lighth****` | `Toouse` | `To the Lighthouse` |
| `The ****** and the ****` | `SoundFury` | `The Sound and the Fury` |

## Power Query M solution

```m
let
  Source = Excel.CurrentWorkbook(){[Name = "Table1"]}[Content], 
  Answer = Table.AddColumn(Source, 
    "Result", 
    each Text.Combine(
      List.Transform(
        List.Zip({Text.Split([Masked String], "*"), Text.ToList([Chars])}), 
        Text.Combine
      )
    )
  )
in
  Answer
```

## Breakdown

The solution adds a new column called `Result` to the source table.

```m
Source = Excel.CurrentWorkbook(){[Name = "Table1"]}[Content]
```

This gets the Excel table named `Table1` from the current workbook.

---

```m
Answer = Table.AddColumn(Source, "Result", each ...)
```

`Table.AddColumn` adds a new calculated column to the source table.

The new column is named `Result`.

For each row, the expression after `each` is evaluated using the values from that row.

---

```m
Text.Split([Masked String], "*")
```

This splits the masked text wherever an asterisk appears.

For example:

```text
*ave *u*!
```

becomes:

| List item | Value |
|---:|---|
| `0` |  |
| `1` | `ave ` |
| `2` | `u` |
| `3` | `!` |

The empty text item appears because there are two asterisks next to each other.

---

```m
Text.ToList([Chars])
```

This converts the replacement characters into a list of individual characters.

For example:

```text
HFn
```

becomes:

| List item | Value |
|---:|---|
| `0` | `H` |
| `1` | `F` |
| `2` | `n` |

---

```m
List.Zip({Text.Split([Masked String], "*"), Text.ToList([Chars])})
```

`List.Zip` pairs the text pieces with the replacement characters.

Using the same example, the two lists are combined like this:

| Pair | Text piece | Replacement character |
|---:|---|---|
| `0` |  | `H` |
| `1` | `ave  ` | `F` |
| `2` |  `u` | `n` |
| `3` | `!` | `null` |

The final text piece has no replacement character because it comes after the final asterisk.

---

```m
List.Transform(
  List.Zip({Text.Split([Masked String], "*"), Text.ToList([Chars])}), 
  Text.Combine
)
```

`List.Transform` applies `Text.Combine` to each pair produced by `List.Zip`.

Each pair is combined into a small text segment:

| Pair | Combined result |
|---:|---|
| `0` | `H` |
| `1` | `ave F` |
| `2` | `un` |
| `3` | `!` |

---

```m
Text.Combine(...)
```

The outer `Text.Combine` combines all the small text segments into the final answer:

```text
Have Fun!
```

---

```m
in
  Answer
```

The query returns the table with the new `Result` column added.

## Why this works

The key idea is to treat the asterisks as split points.

When the masked text is split by `*`, the replacement characters can be inserted between the resulting text pieces.

This avoids writing a loop that searches for each asterisk position manually. Instead, the solution uses list operations:

1. split the masked string into text pieces
2. split the replacement characters into individual characters
3. zip the two lists together
4. combine each pair
5. combine the final result into one text value
6. add that result as a new column

This works even when there are consecutive asterisks, because `Text.Split` preserves the empty text pieces between them.

## Notes

This solution assumes:

- the Excel table is named `Table1`
- the masked text column is named `Masked String`
- the replacement character column is named `Chars`
- the number of characters in `Chars` matches the number of asterisks in `Masked String`

If your table has a different name, update this part:

```m
Excel.CurrentWorkbook(){[Name = "Table1"]}[Content]
```

If your column names are different, update these parts:

```m
[Masked String]
```

and:

```m
[Chars]
```

The solution creates a column named `Result`. If needed, this can be renamed to match the workbook's expected output column.

## Tags

`Power Query` · `M code` · `Text functions` · `Table.AddColumn` · `Text.Split` · `Text.ToList` · `List.Zip` · `List.Transform` · `Text.Combine` · `Excel`
