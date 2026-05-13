# EXCEL CHALLENGE 957 | Character-Driven String Splitting

## Overview

This challenge is about splitting text values using a mix of fixed position rules and character-driven rules.

Each row contains:

- a source text value in `Data`
- one or more split rules in `Split_Rules`
- the expected answer joined with vertical pipes

The goal is to insert split markers before specific characters or positions, then join the result with ` | `.

## Challenge rules

Split the given strings at positions and by the following character-driven split rules:

| Rule | Meaning |
|---|---|
| `@DIGIT` | split before every digit |
| `@UPPER` | split before every uppercase letter |
| `@ALPHA` | split before every alphabetic character |
| number, such as `1` or `5` | split before the character at that zero-based position |

After splitting, join the parts with a vertical pipe:

```text
 |
```

## Source

Challenge creator: [Excel BI (Vijay Verma) on LinkedIn](https://www.linkedin.com/in/excelbi)

Challenge file: [Download practice file](https://lnkd.in/dYBUeV3N)

## Sample data

| Data | Split_Rules | Answer Expected |
|---|---|---|
| `PRODUCT101SALES2023Q4` | `1,8,@DIGIT` | `P \| RODUCT \| 1 \| 0 \| 1SALES \| 2 \| 0 \| 2 \| 3Q \| 4` |
| `USERAdminNorth2024Region` | `@UPPER` | `U \| S \| E \| R \| Admin \| North2024 \| Region` |
| `CAT__DOG___BIRD_FISH` | `(none)` | `CAT__DOG___BIRD_FISH` |
| `2024JANRevenueExpense` | `@ALPHA` | `2024 \| J \| A \| N \| R \| e \| v \| e \| n \| u \| e \| E \| x \| p \| e \| n \| s \| e` |
| `AlphaBetaGammaDelta` | `5,@UPPER` | `Alpha \| Beta \| Gamma \| Delta` |
| `ID999TYPE_A_STATUS_OK` | `1,3` | `I \| D9 \| 99TYPE_A_STATUS_OK` |
| `LondonParisTOKYONewYork` | `@UPPER` | `London \| Paris \| T \| O \| K \| Y \| O \| New \| York` |
| `RedGreenBlueYellowWhite` |  | `RedGreenBlueYellowWhite` |
| `alphaBETA123gamma456` | `@UPPER` | `alpha \| B \| E \| T \| A123gamma456` |

## Power Query M solution

```m
(t as text, rules as nullable text) as text =>
    let
        c = Text.ToList(t),
        r = Text.Split(rules ?? "", ","),
        m = [ #"@DIGIT" = (d) => (try Number.FromText(d) otherwise null) <> null,
            #"@UPPER" = (u) => u = Text.Upper(u) and u <> Text.Lower(u),
            #"@ALPHA" = (a) => Text.Upper(a) <> Text.Lower(a)],
        p = List.RemoveNulls(List.Transform(r, each Record.FieldOrDefault(m, _, null))),
        n = List.RemoveNulls(List.Transform(r, each try Number.FromText(_) otherwise null))
    in
        Text.Combine( List.Transform( List.Positions(c),
            (i) => (if i>0 and (List.Contains(n, i) or List.AnyTrue(List.Transform(p, each _(c{i})))) then " | " else "") & c{i}
        ), ""
    )
```

## How to use the solution

The solution is written as a custom function with two parameters:

| Parameter | Description |
|---|---|
| `t` | the text value to split |
| `rules` | the split rules for that row |

For example, if the function is named `fxSplitText`, it can be used in a custom column like this:

```m
fxSplitText([Data], [Split_Rules])
```

## Breakdown

The solution converts the text into a list of characters, interprets the split rules, then rebuilds the text while inserting ` | ` where a split is needed.

---

```m
(t as text, rules as nullable text) as text =>
```

This defines a custom function.

The function accepts:

- `t`, the text to split
- `rules`, the split rules for that text

The `rules` parameter is nullable, so blank or null rule values can be handled.

---

```m
c = Text.ToList(t)
```

This converts the source text into a list of individual characters.

For example:

```text
ABC123
```

becomes:

| Position | Character |
|---:|---|
| `0` | `A` |
| `1` | `B` |
| `2` | `C` |
| `3` | `1` |
| `4` | `2` |
| `5` | `3` |

Power Query list positions are zero-based.

---

```m
r = Text.Split(rules ?? "", ",")
```

This splits the rule text into a list of individual rules.

The `??` operator replaces a null value with an empty string.

For example:

| Rules | Resulting list |
|---|---|
| `1,8,@DIGIT` | `1`, `8`, `@DIGIT` |
| `5,@UPPER` | `5`, `@UPPER` |
| `@ALPHA` | `@ALPHA` |
| `null` | empty text rule |

---

```m
m = [
    #"@DIGIT" = (d)=> (try Number.FromText(d) otherwise null) <> null,
    #"@UPPER" = (u)=> u = Text.Upper(u) and u <> Text.Lower(u),
    #"@ALPHA" = (a)=> Text.Upper(a) <> Text.Lower(a)
]
```

This creates a record of rule names and their matching functions.

Each field in the record stores a function.

| Rule | Function behavior |
|---|---|
| `@DIGIT` | returns true when the character can be converted to a number |
| `@UPPER` | returns true when the character is uppercase alphabetic text |
| `@ALPHA` | returns true when the character is alphabetic text |

The `@UPPER` test checks two things:

```m
u = Text.Upper(u) and u <> Text.Lower(u)
```

This means the character must already be uppercase, and it must be a letter. The second part prevents digits and symbols from being treated as uppercase.

The `@ALPHA` test checks whether uppercase and lowercase versions of the character are different. If they are different, the character is alphabetic.

---

```m
p = List.RemoveNulls(List.Transform(r, each Record.FieldOrDefault(m, _, null)))
```

This creates a list of active character-driven rule functions.

For each rule in `r`, `Record.FieldOrDefault` looks for a matching field in the rule record `m`.

For example:

| Rule | Match found? |
|---|---|
| `@DIGIT` | digit-checking function |
| `@UPPER` | uppercase-checking function |
| `@ALPHA` | alphabetic-checking function |
| `1` | null |
| `8` | null |

`List.RemoveNulls` removes the rules that are not character-driven functions.

---

```m
n = List.RemoveNulls(List.Transform(r, each try Number.FromText(_) otherwise null))
```

This creates a list of numeric split positions.

For each rule, the solution tries to convert it to a number.

For example:

| Rules | Numeric positions |
|---|---|
| `1,8,@DIGIT` | `1`, `8` |
| `5,@UPPER` | `5` |
| `@ALPHA` | none |

Rules such as `@DIGIT`, `@UPPER`, and `@ALPHA` cannot be converted to numbers, so they become null and are removed.

---

```m
List.Positions(c)
```

This returns the positions of the characters in the text.

For example, if `c` contains six characters, this returns:

```text
0, 1, 2, 3, 4, 5
```

These positions are used to process each character one by one.

---

```m
List.Contains(n, i)
```

This checks whether the current character position `i` is one of the numeric split positions.

For example, if the rules contain `1,8`, the function inserts a split before characters at positions `1` and `8`.

---

```m
List.AnyTrue(List.Transform(p, each _(c{i})))
```

This checks whether any active character-driven rule matches the current character.

For the current character `c{i}`, each rule function in `p` is evaluated.

If at least one function returns true, the expression returns true.

Examples:

| Character | `@DIGIT` | `@UPPER` | `@ALPHA` |
|---|---:|---:|---:|
| `A` | false | true | true |
| `a` | false | false | true |
| `1` | true | false | false |
| `_` | false | false | false |

---

```m
if i > 0 and (List.Contains(n, i) or List.AnyTrue(List.Transform(p, each _(c{i})))) 
then " | " 
else ""
```

This decides whether to insert a split marker before the current character.

A split is inserted when:

- the current position is greater than `0`
- and the position matches a numeric split rule, or the character matches one of the character-driven rules

The `i > 0` condition prevents the result from starting with a leading pipe.

---

```m
& c{i}
```

This appends the current character after the optional split marker.

So each character becomes either:

```text
character
```

or:

```text
 | character
```

depending on whether a split is needed.

---

```m
Text.Combine(..., "")
```

Finally, all generated text fragments are combined into one result.

The delimiter is an empty string because the pipe markers have already been inserted as part of the character transformation.

## Why this works

The key idea is to separate the rules into two groups:

1. numeric position rules
2. character-driven rules

Numeric rules are stored as positions in `n`.

Character-driven rules are stored as functions in `p`.

Then the solution checks each character once. If the character position or character type matches one of the rules, it inserts ` | ` before that character.

This makes the function flexible because multiple rule types can be combined, such as:

```text
1,8,@DIGIT
```

or:

```text
5,@UPPER
```

## Notes

This solution assumes:

- the source text is provided as the first argument
- the split rules are provided as the second argument
- numeric split positions use zero-based indexing
- rule names match the expected values exactly, such as `@DIGIT`, `@UPPER`, and `@ALPHA`

If the rules are stored in a table with columns named `Data` and `Split_Rules`, create a custom column using the function, for example:

```m
fxSplitText([Data], [Split_Rules])
```

## Tags

`Power Query` · `M code` · `Custom function` · `Text.ToList` · `List.Transform` · `Record.FieldOrDefault` · `List.Positions` · `Text.Combine` · `Excel`
