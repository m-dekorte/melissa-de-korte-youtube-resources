# Challenge 112 | Split First Name(s) and Other Names

## Overview

This challenge is about splitting a full name into two parts:

- `First Name`
- `Other Names`

The first name portion can contain one or more words, and those first name words are written in all uppercase letters.

The solution needs to be dynamic, so it should not assume that the first name always contains only one word.

## Challenge rules

Split each value in the `Name` column using these rules:

1. The first name part is written in all uppercase.
2. The first name part may contain one or more words.
3. All remaining words should be returned as other names.
4. The solution must be dynamic.

## Source

Challenge creator: [Crispo Mwangi on LinkedIn](https://www.linkedin.com/in/crispo-mwangi-6ab49453)

Excel file: [Download workbook](https://lnkd.in/dMf4C5KK)

## Sample structure

The source table is expected to contain a column named `Name`.

Example pattern:

| Name | First Name | Other Names |
|---|---|---|
| `JOHN Smith` | `JOHN` | `Smith` |
| `MARY ANN Johnson` | `MARY ANN` | `Johnson` |
| `PETER JAMES van der Merwe` | `PETER JAMES` | `van der Merwe` |
| `ANNA Maria Lopez` | `ANNA` | `Maria Lopez` |

## Power Query M solution

```m
Table.FromRows(
    List.Transform(
        Source[Name],
        each [
            a = Text.Split(_, " "),
            p = List.FirstN(a, each Text.Remove(_, {"A" .. "Z"}) = ""),
            r = {Text.Combine(p, " "), Text.Combine(List.Skip(a, List.Count(p)), " ")}
        ][r]
    ),
    {"First Name", "Other Names"}
)
```

## Breakdown

The solution loops through the `Name` column, identifies the consecutive uppercase words at the start of each name, and returns a two-column table.

---

```m
Source[Name]
```

This extracts the values from the `Name` column as a list.

Each item in the list is one full name to process.

---

```m
List.Transform(
    Source[Name],
    each ...
)
```

`List.Transform` loops through each name in the list.

For every name, it returns a two-item list:

1. the first name part
2. the remaining names

Those two values later become the columns `First Name` and `Other Names`.

---

```m
a = Text.Split(_, " ")
```

This splits the current name into separate words using a space as the delimiter.

For example:

```text
MARY ANN Johnson
```

becomes:

| List item | Value |
|---:|---|
| `0` | `MARY` |
| `1` | `ANN` |
| `2` | `Johnson` |

The underscore `_` represents the current name being processed.

---

```m
p = List.FirstN(a, each Text.Remove(_, {"A" .. "Z"}) = "")
```

This gets the first consecutive words that are written entirely in uppercase letters.

`Text.Remove(_, {"A" .. "Z"})` removes all uppercase letters from the current word.

If nothing is left after removing uppercase letters, the word was made only of uppercase letters.

Example:

| Word | After removing `A` to `Z` | Uppercase first-name word? |
|---|---|---|
| `MARY` | `""` | yes |
| `ANN` | `""` | yes |
| `Johnson` | `ohnson` | no |

`List.FirstN` keeps taking words from the start of the list while the condition is true.

So for:

```text
MARY ANN Johnson
```

the list `p` becomes:

```text
MARY, ANN
```

---

```m
Text.Combine(p, " ")
```

This combines the uppercase first name words back into a single text value.

For example:

```text
MARY, ANN
```

becomes:

```text
MARY ANN
```

---

```m
List.Skip(a, List.Count(p))
```

This skips the first name words and keeps the remaining words.

For example, if:

```text
a = MARY, ANN, Johnson
```

and:

```text
p = MARY, ANN
```

then `List.Count(p)` is `2`, so the first two words are skipped.

The remaining list is:

```text
Johnson
```

---

```m
Text.Combine(List.Skip(a, List.Count(p)), " ")
```

This combines the remaining words into the `Other Names` value.

For example:

```text
Johnson
```

or:

```text
van der Merwe
```

---

```m
r = {Text.Combine(p, " "), Text.Combine(List.Skip(a, List.Count(p)), " ")}
```

This creates the final two-item list for the current row.

The first item becomes `First Name`.

The second item becomes `Other Names`.

---

```m
[a = ..., p = ..., r = ...][r]
```

This creates a temporary record with three fields:

| Field | Purpose |
|---|---|
| `a` | stores the name split into words |
| `p` | stores the uppercase first name words |
| `r` | stores the final two values to return |

The final `[r]` returns only the result list.

This is a compact way to use intermediate steps without writing a separate `let ... in` expression inside the function.

---

```m
Table.FromRows(
    ...,
    {"First Name", "Other Names"}
)
```

`Table.FromRows` converts the list of row values into a table.

The second argument provides the output column names:

- `First Name`
- `Other Names`

## Why this works

The key idea is to identify the first name part by checking words from the left until the first non-uppercase word appears.

This makes the solution dynamic because it works whether the first name contains:

- one uppercase word
- two uppercase words
- several uppercase words

It does not need a fixed number of first-name words.

## Notes

This solution assumes:

- the previous step is named `Source`
- the source table contains a column named `Name`
- first name words contain only uppercase letters from `A` to `Z`
- first name words appear at the start of the full name

If your source column has a different name, update this part:

```m
Source[Name]
```

If your previous step has a different name, update this part:

```m
Source
```

## Tags

`Power Query` · `M code` · `Text.Split` · `Text.Remove` · `List.FirstN` · `List.Skip` · `Table.FromRows` · `Excel`
