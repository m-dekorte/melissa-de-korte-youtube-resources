# Excel Challenge 955: Longest Valid Parentheses Substring

## Overview

This challenge is about finding the length of the longest valid, well-formed parentheses substring in a text value.

A valid parentheses substring is one where every opening parenthesis `(` has a matching closing parenthesis `)` in the correct order, and all matching pairs are properly nested.

## Challenge rules

Find the length of the longest valid parentheses substring.

A valid parentheses substring must follow these rules:

1. Every `(` has a matching `)`.
2. Every `)` has a matching `(` that appears before it.
3. Matching pairs must be properly nested.
4. Return the length of the longest valid substring.

## Source

Challenge creator: [Excel BI (Vijay Verma) on LinkedIn](https://www.linkedin.com/in/excelbi)

Excel file: [Download practice file](https://lnkd.in/dduxcTSG)

## Examples

| Input | Longest valid substring | Answer |
|---|---|---:|
| `()((())()` | `(())()` | `6` |
| `())()())(` | `()()` | `4` |

## Power Query M solution

```m
(s as text) as number =>
    [
        c = Text.ToList(s), 
        n = List.Count(c), 
        f = (i, l, m) =>
            if i = n then m
            else if c{i} = "(" then @f(i + 1, {i} & l, m)
            else [ t = List.Skip(l, 1), 
                   r = if List.IsEmpty(t) then @f(i + 1, {i}, m)
                      else @f(i + 1, t, List.Max({m, i - t{0}}))
                ][r], 
        a = f(0, {- 1}, 0)
    ][a]
```

## How to use the solution

The solution is written as a custom function with one parameter:

| Parameter | Description |
|---|---|
| `s` | the text value containing parentheses |

For example, if the function is named `fxLongestValidParentheses`, it can be used in a custom column like this:

```m
fxLongestValidParentheses([String])
```

## Breakdown

The solution scans the text from left to right and keeps track of positions that help calculate the length of valid parentheses substrings.

It uses a recursive function to process each character.

---

```m
c = Text.ToList(s)
```

This converts the text into a list of individual characters.

For example:

```text
())()())(
```

becomes a list of characters:

| Position | Character |
|---:|---|
| `0` | `(` |
| `1` | `)` |
| `2` | `)` |
| `3` | `(` |
| `4` | `)` |
| `5` | `(` |
| `6` | `)` |
| `7` | `)` |
| `8` | `(` |

Power Query uses zero-based positions, so the first character is at position `0`.

---

```m
n = List.Count(c)
```

This stores the total number of characters in the text.

The recursive function uses this value to know when it has reached the end.

---

```m
f = (i, l, m) =>
```

This defines a recursive helper function named `f`.

The function uses three parameters:

| Parameter | Purpose |
|---|---|
| `i` | current character position |
| `l` | list of tracked positions |
| `m` | maximum valid length found so far |

The list `l` works like a stack of positions.

It stores unmatched opening parentheses and the most recent boundary where a valid substring cannot cross.

---

```m
if i = n then m
```

This is the stopping condition.

When the current position `i` reaches the total number of characters `n`, the scan is complete, so the function returns `m`, the longest valid length found.

---

```m
else if c{i} = "(" then @f(i + 1, {i} & l, m)
```

If the current character is an opening parenthesis `(`, its position is added to the front of the list `l`.

The function then moves to the next character.

The `@` symbol is used because the function calls itself recursively.

---

```m
else [ 
    t = List.Skip(l, 1), 
    r = ...
][r]
```

If the current character is not an opening parenthesis, it is treated as a closing parenthesis `)`.

The solution removes the first item from `l` using `List.Skip(l, 1)`.

This is the equivalent of matching the current closing parenthesis with the most recent unmatched opening parenthesis.

---

```m
r = if List.IsEmpty(t) then @f(i + 1, {i}, m)
```

If removing the first item leaves the list empty, it means there is no valid opening parenthesis available to match the current closing parenthesis.

In that case, the current position becomes the new boundary.

That boundary is stored as `{i}`.

A valid substring cannot start before this unmatched closing parenthesis.

---

```m
else @f(i + 1, t, List.Max({m, i - t{0}}))
```

If the list is not empty, a valid substring has been found or extended.

The length of the current valid substring is calculated as:

```m
i - t{0}
```

Here:

- `i` is the current closing parenthesis position
- `t{0}` is the most recent boundary or unmatched opening parenthesis position after the match

The maximum length is updated using:

```m
List.Max({m, i - t{0}})
```

This keeps the larger value between:

- the previous maximum length
- the current valid substring length

---

```m
a = f(0, {- 1}, 0)
```

This starts the recursive scan.

The initial values are:

| Argument | Initial value | Meaning |
|---|---:|---|
| `i` | `0` | start at the first character |
| `l` | `{-1}` | initial boundary before the text starts |
| `m` | `0` | no valid substring found yet |

The initial boundary `-1` allows a valid substring starting at position `0` to be measured correctly.

For example, if the first two characters are `()`, then at position `1` the length is:

```text
1 - (-1) = 2
```

---

```m
[a]
```

The final `[a]` returns the answer from the temporary record.

## Why this works

The key idea is to track positions, not just counts.

A simple count of opening and closing parentheses is not enough, because the pairs must appear in the correct order and be properly nested.

The list `l` stores the positions needed to calculate valid substring lengths:

- unmatched opening parentheses
- the latest invalid boundary

When a valid closing parenthesis is found, the distance between the current position and the latest boundary gives the length of the current valid substring.

The maximum of these lengths is the final answer.

## Step example

For this input:

```text
())()())(
```

The longest valid substring is:

```text
()()
```

The answer is:

```text
4
```

The unmatched `)` characters reset the boundary, so the function only measures valid substrings after the latest invalid position.

## Notes

This solution assumes:

- the input contains parentheses characters
- the function receives a text value
- the result should be the length of the longest valid parentheses substring

If the source column is named `String`, use the function in a custom column like this:

```m
fxLongestValidParentheses([String])
```

## Tags

`Power Query` · `M code` · `Custom function` · `Recursion` · `Text.ToList` · `List.Skip` · `List.Max` · `Parentheses` · `Excel`
