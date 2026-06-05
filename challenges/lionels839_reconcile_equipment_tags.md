# Reconcile Free-Text Equipment Names with Official Tags

**Date:** 2026-06-05  
**Requester:** @lionels839  
**Requested on Video:** [Fixed-Width Text Files in Power Query: a Dynamic Solution](https://www.youtube.com/watch?v=I6Ld9Zn9IR4)  
**Solution Video:** [Check out my Channel](https://www.youtube.com/@melissa_de_korte/)

## Challenge

I work in the oil industry and have a list of tens of thousands of equipment tags such as:

```text
011P0001A
012P0001
013P0008A
13P0008B
```

The official equipment tag follows this general structure:

```text
L UU T NNNN S
```

Where:

| Part | Meaning | Example |
|---|---|---|
| `L` | Location, 1 digit | `0` |
| `UU` | Unit, 2 digits | `11` |
| `T` | Equipment type, 1 letter | `P` |
| `NNNN` | Equipment number, 4 digits | `0001` |
| `S` | Optional spare suffix | `A` or `B` |

For example:

```text
011P0001
```

Means:

- location: `0`
- unit: `11`
- equipment type: `P`
- equipment number: `0001`
- optional spare: *blank*

## Problem

Maintenance requests are created using free-text descriptions. Users do not always type the official equipment tag.

For example, the official tag:

```text
011P0001
```

May appear as any of the following:

```text
011P001
011P01
011P1
11P0001
11P001
11P1
```

Or even as below, giving context like « *the pump P1 of unit 11 is broken* »
```text
P0001
P001
P01
P1
```

From this sentence, the expected reconstructed equipment tag is:

```text
011P0001
```

## Goal

Create a Power Query solution that can extract and normalize equipment tags from free-text maintenance descriptions, so that the maintenance data can be reconciled with the official equipment master data.

## Typical Raw Data

The raw data usually contains columns such as:

| Column | Description |
|---|---|
| `tag` | Free-text or partial equipment tag |
| `description of the problem` | Maintenance description |
| `Start date of maintenance` | Maintenance start date |
| `End date of maintenance` | Maintenance end date |

## Power Query M Solution

### Function: `fxGetEquipmentTag`

```m
(s as text, optional defaultLocation as text) as text =>
let 
    words = {"LOCATION", "UNIT", "PUMP", "TYPE", "NUMBER", "SPARE"},
    wordMap = {1, 2, 3, 3, 4, 5},
    splitChars = " ,.;:!?",
    location = defaultLocation ?? "0",
    input = Text.Upper(Text.Trim(s)),

    getTag = (t) => let
        a = List.Combine(
            List.Transform(
                Splitter.SplitTextByCharacterTransition(each true, {"A".."Z"})(t),
                each Splitter.SplitTextByCharacterTransition({"A".."Z"}, each true)(_)
            )
        ), 
        b = if List.Count(a) > 2 
            then Text.PadStart(a{0}, 3, location) & a{1} & Text.PadStart(a{2}, 4, "0") & (a{3}? ?? "")
            else t
    in b,
    extractTag = (t) => let
        a = List.Select(Text.SplitAny(t, splitChars), each (_ ?? "") <> ""),
        b = List.RemoveNulls(
            List.Transform(List.Positions(a), each [
                pos = List.PositionOf(words, a{_}),
                out = if pos >= 0 and _ < List.Count(a) - 1 
                    then {wordMap{pos}, a{_ + 1}}
                    else null
                ][out]
            )
        ),
        c = Text.Combine(List.Transform(List.Sort(b, each _{0}),each _{1}), "")
    in c,
    result = if Text.Contains(input, " ") then getTag(extractTag(input)) else getTag(input)
in 
    result
```

## Demo Query

```m
let
    Source = Table.FromValue({
        "011P001",
        "011P01",
        "011P1",
        "11P0001",
        "11P001",
        "11P1",
        "011P0001A",
        "012P0001",
        "013P0008A",
        "13P0008B",
        "blabla bla the pump P1 of unit 11 is broken"
    }),
    getTag = Table.AddColumn(Source, "test", each fxGetEquipmentTag([Value]))
in
    getTag
```

## Example Output

| Input | Expected normalized tag |
|---|---|
| `011P001` | `011P0001` |
| `011P01` | `011P0001` |
| `011P1` | `011P0001` |
| `11P0001` | `011P0001` |
| `11P001` | `011P0001` |
| `11P1` | `011P0001` |
| `011P0001A` | `011P0001A` |
| `012P0001` | `012P0001` |
| `013P0008A` | `013P0008A` |
| `13P0008B` | `013P0008B` |
| `blabla bla the pump P1 of unit 11 is broken` | `011P0001` |

## Explanation of the Word Mapping

The context-word extraction uses this mapping:

```m
words = {"LOCATION", "UNIT", "PUMP", "TYPE", "NUMBER", "SPARE"}
wordMap = {1, 2, 3, 3, 4, 5}
```

This means:

| Word | Sort order | Meaning |
|---|---:|---|
| `LOCATION` | 1 | Location prefix |
| `UNIT` | 2 | Unit number |
| `PUMP` | 3 | Equipment type |
| `TYPE` | 3 | Equipment type |
| `NUMBER` | 4 | Equipment number |
| `SPARE` | 5 | Optional spare suffix |

For example:

```text
blabla bla the pump P1 of unit 11 is broken
```

The function extracts:

| Found word | Extracted value | Sort order |
|---|---|---:|
| `PUMP` | `P1` | 3 |
| `UNIT` | `11` | 2 |

After sorting by the mapping, the extracted values become:

```text
11P1
```

Then `getTag` normalizes this into:

```text
011P0001
```

## Notes

- fxGetEquipmentTag takes a single text input, if needed, input values can be concatenated.
- The default location is `0`, but another location prefix can be supplied through the optional `defaultLocation` parameter.
- The solution converts the input to uppercase, trims leading and trailing spaces, before processing.
- The current split characters are: `" ,.;:!?"`
- Additional delimiters can be added such as `/`, `-`, `_`, parentheses, or line breaks.
- This function is designed as a practical reconciliation helper and should be validated against the official equipment master list after normalization.
