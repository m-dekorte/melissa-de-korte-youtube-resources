# Function.From: From Fixed Parameters to Flexible Functions

You need your Power Query custom function to accept more inputs. So you add another parameter. Then another. Then another. But you're only moving the limit.

This video explores different approaches — using optional parameters, a list input, and finally `Function.From` — to build functions that accept a varying number of arguments.

Video: [Watch on YouTube](https://youtu.be/Whexruj9ZQg)

LinkedIn: [Connect with Melissa](https://www.linkedin.com/in/melissa-de-korte/)

---

## The Problem

A custom function with fixed parameters can only accept a set number of inputs.

```m
// addNumbers
(a as number, b as number) as number =>
    a + b
```

This function takes exactly two arguments. No more.

If you need to add a third value, you have to create a third parameter.  
And a fourth. And a fifth. Each time the requirements change, the function signature changes.

The question is: how do you make this flexible?

---

## Approach 1: Optional Parameters

Adding optional parameters creates some room.

```m
// addMoreNumbers
(a as number, b as number, 
   optional c as nullable number, 
   optional d as nullable number, 
   optional e as nullable number) as number =>
   a + b + (c ?? 0) + (d ?? 0) + (e ?? 0)
```

When an optional argument is not supplied, its value is `null`. Adding `null` to a number propagates the `null`, so each optional parameter uses coalesce (`??`) to replace `null` with `0`.

The coalesce operator `??` evaluates the expression on its left. If that expression is `null`, it returns the value on the right instead.

```m
(c ?? 0)
```

If `c` is `null`, this evaluates to `0`. The parentheses ensure the coalesce expression is evaluated before the addition.

This approach works, but it moves the limit rather than removing it. You can handle five inputs today, but not six.

---

## Approach 2: A Single List Parameter

Passing all values in a list removes the limit entirely.

```m
// addMore
(values as list) as number =>
    List.Sum(values)
```

`List.Sum` handles `null` values automatically, so no additional logic is needed.

```m
addMore({1, 2, 3, 4, 5, null})
```

Six values, no problem. Add more at any time.

But there is one drawbacks.

The user must know how to construct and pass a list. The input must be wrapped in `{ }`, which is less intuitive than passing separate arguments.

---

## Approach 3: Function.From

`Function.From` creates a function that accepts a variable number of arguments, collects them into a list, and passes that list to a processing function.

It takes two arguments:

| Argument | Purpose |
|---|---|
| A **function type** | Describes the function, parameter type, and return type |
| A **function** | Receives the collected list and produces a result |

### Making Function.From Visible

To understand what happens inside `Function.From`, return the collected input unchanged using `each _`:

```m
// Explore
Function.From(
    type function (value as nullable number) as any,
    each _
)
```

Invoking this function reveals the mechanism:

```m
Explore(10)          // returns {10}
Explore(10, 20, 30)  // returns {10, 20, 30}
Explore(10, null, 5) // returns {10, null, 5}
```

Even though the function type declares a single parameter `value`, the function accepts any number of arguments.  
All arguments are collected in a list and handed to the second argument — the processing function.

### Using Function.From with List.Sum

Replace `each _` with `List.Sum` to create a function that sums any number of inputs:

```m
// addUnlimited
Function.From(
    type function (value as nullable number) as nullable number,
    List.Sum
)
```

```m
addUnlimited(10, 20, 30, 40, 50, null)
```

This function:

- accepts any number of arguments
- does not require the inputs to be passed in a list

---

## How Function.From Works

The mechanism is straightforward:

```
Separate arguments go in
    → collected in a list
        → list is handed to the processing function
            → processing function produces a result
```

---

## All Code Samples

### addNumbers

```m
(a as number, b as number) as number =>
    a + b
```

### addMoreNumbers

```m
(a as number, b as number, 
   optional c as nullable number, 
   optional d as nullable number, 
   optional e as nullable number) as number =>
   a + b + (c ?? 0) + (d ?? 0) + (e ?? 0)
```

### addMore

```m
(values as list) as number =>
    List.Sum(values)
```

### myList

```m
let
    Source = Table.FromValue({1, 2, 3, 4, 5, 6})
in
    Source
```

### Explore

```m
Function.From(
    type function (value as nullable number) as any,
    each _
)
```

### addUnlimited

```m
Function.From(
    type function (value as nullable number) as nullable number,
    List.Sum
)
```

---

## Quick Reference

| Approach | Flexible? | Syntax |
|---|---|---|
| Fixed parameters | No | `fx(1, 2)` |
| Optional parameters | Limited |`fx(1, 2, 3)` |
| List parameter | Yes | `fx({1, 2, 3})` |
| `Function.From` | Yes | `fx(1, 2, 3)` |

---

## Key Functions

| Function | Purpose |
|---|---|
| `Function.From` | Creates a function that collects arguments into a list |
| `List.Sum` | Sums non-null values in a list; returns `null` when all values are null |

---

## Resources

- [Function.From — Microsoft documentation](https://learn.microsoft.com/en-us/powerquery-m/function-from?wt.mc_id=MVP_371969)
- [List.Sum — Microsoft documentation](https://learn.microsoft.com/en-us/powerquery-m/list-sum?wt.mc_id=MVP_371969)

---

## Tags

`Power Query` · `M code` · `Function.From` · `Custom function` · `Function type` · `List.Sum` · `Optional parameters` · `Type system` · `Power BI` · `Excel`
