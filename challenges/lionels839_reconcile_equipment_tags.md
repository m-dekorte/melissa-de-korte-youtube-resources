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
let
    fxGetEquipmentTag =
    let
        fnGetEquipmentTag = (s as text, optional defaultLocation as nullable text, optional overwrite as logical) as text =>
        let 
            input = Text.Upper(Text.Trim(s)),
            splitChars = " ,.;:!?",
            location = defaultLocation ?? "0",
            choice = overwrite ?? true,
            wordMap = [LOCATION = "s1", UNIT = "s2", PUMP = "s3", TYPE = "s3", NUMBER = "s4", SPARE = "s5" ],

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
                b = List.Accumulate(
                    a,
                    [pending = null, r = [s1 = "", s2 = "", s3 = "", s4 = "", s5 = "" ]],
                    (x, y) => [
                        r = if x[pending] = null then x[r] else if choice then
                            x[r] & Record.FromList({y}, {x[pending]}) else
                            x[r] & Record.FromList({Record.Field(x[r], x[pending]) & y},{x[pending]}),
                        pending = Record.FieldOrDefault(wordMap, y, null)
                    ]
                ),
                outcome = Text.Combine(Record.FieldValues(b[r]), "")
            in outcome,
            result = if Text.Contains(input, " ") then getTag(extractTag(input)) else getTag(input)
        in result,

        stringParameter = type text meta [
            Documentation.FieldCaption = "s",
            Documentation.FieldDescription = "The source text to convert into a standardized tag. The text may already be a compact tag, or it may contain keywords such as Location, Unit, Pump, Type, Number, and Spare.",
            Documentation.SampleValues = {"12A34", "Location 12 Unit A Pump 34"}
        ],
        defaultLocationParameter = type nullable text meta[
            Documentation.FieldCaption = "defaultLocation",
            Documentation.FieldDescription = "Optional padding character used when the location part of the tag is shorter than 3 characters. When omitted, 0 is used.",
            Documentation.SampleValues = { "0", "9" }
        ],
        overwriteParameter = type logical meta[
            Documentation.FieldCaption = "overwrite",
            Documentation.FieldDescription = "When true, repeated keyword values overwrite earlier values. When false, repeated values for the same keyword position are concatenated. When omitted, overwrite is used.",
            Documentation.SampleValues = {true, false}
        ],
        addParameterTypes = type function (
            s as stringParameter,
            optional defaultLocation as defaultLocationParameter,
            optional overwrite as overwriteParameter
        ) as text meta [
            Documentation.Name = "fxGetEquipmentTag",
            Documentation.Description = "Converts a text value into a standardized tag by extracting or formatting location, unit, pump/type, number, and spare components.",
            Documentation.Version = "1.00: Initial Version",
            Documentation.Author = "Melissa de Korte",
            Documentation.Examples = {
                [
                    Description = "Format a compact tag",
                    Code = "fxGetEquipmentTag(""11P1"")",
                    Result = "011P0001"
                ],
                [
                    Description = "Extract a tag from labelled text",
                    Code = "fxGetEquipmentTag(""the pump P1 of unit 11 is broken"")",
                    Result = "011P0001"
                ],
                [
                    Description = "Use a different location padding character",
                    Code = "fxGetEquipmentTag(""11P1"", ""9"")",
                    Result = "911P0001"
                ],
                [
                    Description = "Concatenate repeated keyword values instead of overwriting them",
                    Code = "fxGetEquipmentTag(""Location 1 Location 2 the pump P1 of unit 11 is broken"", ""0"", false)",
                    Result = "1211P0001"
                ]
            }
        ]
    in
        Value.ReplaceType(fnGetEquipmentTag, addParameterTypes)
in
    fxGetEquipmentTag
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
wordMap = [LOCATION = "s1", UNIT = "s2", PUMP = "s3", TYPE = "s3", NUMBER = "s4", SPARE = "s5" ]
```

This means:

| Word | Position | Meaning |
|---|---:|---|
| `LOCATION` | s1 | Location prefix |
| `UNIT` | s2 | Unit number |
| `PUMP` | s3 | Equipment type |
| `TYPE` | s3 | Equipment type |
| `NUMBER` | s4 | Equipment number |
| `SPARE` | s5 | Optional spare suffix |

For example:

```text
blabla bla the pump P1 of unit 11 is broken
```

The function extracts:

| Found word | Extracted value | Position |
|---|---|---:|
| `PUMP` | `P1` | s3 |
| `UNIT` | `11` | s2 |

After the record field values are mapped, the extracted values become:

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
- The default choice is `overwrite` is `true`. To append repeated keyWord values, set the optional `overwrite` parameter to `false`.
- The solution converts the input to uppercase, trims leading and trailing spaces, before processing.
- The current split characters are: `" ,.;:!?"`
- Additional delimiters can be added such as `/`, `-`, `_`, parentheses, or line breaks.
- This function is designed as a practical reconciliation helper and should be validated against the official equipment master list after normalization.
