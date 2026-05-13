# Remove Consecutive Duplicate Characters

## Overview

This challenge is about removing repeated characters when they appear consecutively in a text value.

If the same character appears multiple times in a row, only the first occurrence should be kept.

## Challenge rules

If a character appears consecutively, remove all consecutive repeated characters except the first one.

For example:

```text
xxxyxyyz
```

In this string:

- `x` appears 3 times consecutively at the start, so only the first `x` is kept
- `yy` appears 2 times consecutively later in the string, so only the first `y` is kept

Expected answer:

```text
xyxyz
```

## Source

Challenge creator: [Meganathan Elumalai on LinkedIn](https://www.linkedin.com/in/meganathan-elumalai-a3449221/)

Excel file: [Download workbook](https://lnkd.in/gh8BrPAK)

## Sample data

| String | Answer Expected |
|---|---|
| `abb` | `ab` |
| `abccabca` | `abcabca` |
| `aabbbccaa` | `abca` |
| `uniiiveersse` | `universe` |
| `Eexxxxxccceeell` | `Excel` |

## Power Query M solution

```m
(s as nullable text) as text =>
    [
        x = s ?? "", 
        a = Text.Start(x, 1)
            & Text.Combine(
                List.Transform({1 .. Text.Length(x) - 1}, 
                    each if Comparer.OrdinalIgnoreCase(Text.At(x, _ - 1), Text.At(x, _)) = 0 
                    then "" else Text.At(x, _)
                )
            )
    ][a]
```

## How to use the solution

The solution is written as a custom function with one parameter:

| Parameter | Description |
|---|---|
| `s` | the text value to clean |

For example, if the function is named `fxRemoveConsecutiveDuplicates`, it can be used in a custom column like this:

```m
fxRemoveConsecutiveDuplicates([String])
```

## Breakdown

The solution compares each character with the character immediately before it.

If the current character is the same as the previous character, it is removed. If it is different, it is kept.

---

```m
(s as nullable text) as text =>
```

This defines a custom function.

The function accepts one value, `s`, which can be text or null, and returns a text value.

---

```m
x = s ?? ""
```

This handles null values.

If `s` is null, it is replaced with an empty string.

If `s` already contains text, that text is used.

---

```m
Text.Start(x, 1)
```

This keeps the first character of the text.

The first character is always retained because there is no previous character to compare it with.

For example:

| Input | First character kept |
|---|---|
| `abb` | `a` |
| `uniiiveersse` | `u` |
| `Eexxxxxccceeell` | `E` |

---

```m
List.Transform({1 .. Text.Length(x) - 1}, each ...)
```

This loops through the remaining character positions in the text, starting from position `1`.

Power Query uses zero-based positions, so:

| Position | Meaning |
|---:|---|
| `0` | first character |
| `1` | second character |
| `2` | third character |

The first character is handled separately by `Text.Start(x, 1)`, so the loop starts at the second character.

---

```m
Text.At(x, _ - 1)
```

This gets the previous character.

---

```m
Text.At(x, _)
```

This gets the current character.

---

```m
Comparer.OrdinalIgnoreCase(Text.At(x, _ - 1), Text.At(x, _)) = 0
```

This compares the previous character with the current character.

`Comparer.OrdinalIgnoreCase` performs a case-insensitive comparison.

That means characters such as `E` and `e` are treated as equal.

For example, in:

```text
Eexxxxxccceeell
```

the initial `E` and `e` are treated as consecutive duplicates, so the result starts with:

```text
E
```

not:

```text
Ee
```

---

```m
then "" else Text.At(x, _)
```

If the current character is the same as the previous character, the function returns an empty string.

If the current character is different from the previous character, the current character is kept.

---

```m
Text.Combine(...)
```

This combines all retained characters from the loop into one text value.

---

```m
Text.Start(x, 1) & Text.Combine(...)
```

This joins the first character with the cleaned remaining characters.

Together, these parts produce the final text with consecutive duplicates removed.

---

```m
[
    x = ...,
    a = ...
][a]
```

This creates a temporary record with two fields:

| Field | Purpose |
|---|---|
| `x` | stores the cleaned input value, replacing null with empty text |
| `a` | stores the final answer |

The final `[a]` returns only the result.

This is a compact way to use intermediate steps without writing a separate `let ... in` expression.

## Why this works

The key idea is that a character should only be removed when it is the same as the character immediately before it.

This means the solution removes consecutive duplicates but keeps repeated characters when they are separated by other characters.

For example:

```text
xyxyz
```

The character `x` appears twice, but not consecutively. Both `x` values are kept.

## Notes

This solution assumes:

- the input column contains text values or nulls
- the goal is to compare characters case-insensitively
- only consecutive repeated characters should be removed

If the source column is named `String`, use the function in a custom column like this:

```m
fxRemoveConsecutiveDuplicates([String])
```

## Tags

`Power Query` · `M code` · `Custom function` · `Text.Start` · `Text.At` · `Text.Combine` · `List.Transform` · `Comparer.OrdinalIgnoreCase` · `Excel`
