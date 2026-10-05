# Excel-style calculations and rounding

This custom Power Query M function is a **configurable calculation factory**. You specify **what to do** and optionally **how to do it** with an options record. The function then returns a configured function that accepts the values to calculate.

The aim is **Excel-style calculations and rounding**, while retaining Power Query's native value conversion, null handling, errors, and numeric behavior.

The design supports familiar calculations such as `SUM`, `AVERAGE`, `PRODUCT`, `POWER`, `SQRT`, `FACT`, `LOG`, and `MOD`, together with Excel-style `ROUND`, `ROUNDUP`, and `ROUNDDOWN`. Values can optionally be rounded to a specified number of digits or to a multiple such as `0.05`, `5`, or `100`.

---

## Function

```m
let
    fxExcelStyleCalc = (calculation as text, optional options as nullable record) as function =>
    let
        get = (field as text, default as any) as any =>
            Record.FieldOrDefault(options ?? [], field, default),
        f = (fn as function) as function =>
            Function.From(
                type function (value as nullable number) as nullable number,
                fn
            ),
        fold = (fn as function) as function =>
            f(each List.Accumulate(List.Skip(_), _{0}, fn)),
        calculations = [
            SUM = f(List.Sum),
            AVERAGE = f(List.Average),
            PRODUCT = f(List.Product),
            MIN = f(List.Min),
            MAX = f(List.Max),
            MEDIAN = f(List.Median),
            SUBTRACT = fold(Value.Subtract),
            DIVIDE = fold(Value.Divide),
            POWER = Number.Power,
            SQRT = Number.Sqrt,
            EXP = Number.Exp,
            FACT = each Number.Factorial(Number.RoundTowardZero(_)),
            ABS = Number.Abs,
            SIGN = Number.Sign,
            LN = Number.Ln,
            LOG10 = Number.Log10,
            LOG = (n, optional b) => Number.Log(n, b ?? 10),
            MOD = (n, d) => n - d * Number.RoundDown(n / d)
        ],
        roundingType = [
            ROUND = (n, d) => Number.Round(n, d, RoundingMode.AwayFromZero),
            ROUNDUP = Number.RoundAwayFromZero,
            ROUNDDOWN = Number.RoundTowardZero
        ],
        calc = Record.Field(
            calculations,
            Text.Upper(calculation)
        ),
        rnd = Record.Field(
            roundingType,
            Text.Upper(get("Rounding", "ROUND"))
        ),
        ignoreNull = get("IgnoreNull", false),
        culture = get("Culture", ""),
        digits = get("Digits", null),
        step = get("Step", null),
        result = Function.From(
            type function (value as any) as nullable number,
            (arguments as list) as nullable number => [
                converted = List.Transform(
                    arguments,
                    each Number.From(_, culture)
                ),
                values = if ignoreNull
                    then List.RemoveNulls(converted)
                    else converted,
                calculated = if List.IsEmpty(values)
                    or (not ignoreNull
                        and List.NonNullCount(values) <> List.Count(values))
                    then null
                    else Function.Invoke(calc, values),
                rounded = if calculated = null
                    then calculated
                    else if step <> null
                    then [
                        unit = Number.Abs(Number.From(step, culture)),
                        result = if unit = 0
                            then error "Step cannot be zero."
                            else Value.Multiply(
                                rnd(
                                    Value.Divide(calculated, unit, Precision.Decimal), 0
                                ), unit, Precision.Decimal
                            )
                    ][result]
                    else if digits <> null
                    then [
                        value = Number.From(digits, culture),
                        result = if value <> Number.RoundTowardZero(value)
                            then error "Digits must be a whole number."
                            else rnd(calculated, value)
                    ][result]
                    else calculated
            ][rounded]
        )
    in
        result,
    availableCalculations = {
        "SUM",
        "AVERAGE",
        "PRODUCT",
        "MIN",
        "MAX",
        "MEDIAN",
        "SUBTRACT",
        "DIVIDE",
        "POWER",
        "SQRT",
        "EXP",
        "FACT",
        "ABS",
        "SIGN",
        "LN",
        "LOG10",
        "LOG",
        "MOD"
    },
    fxWithDocumentation =
        let
            calculationParameter = type text
                meta [
                    Documentation.FieldCaption = "Calculation",
                    Documentation.FieldDescription =
                        "Required. Selects the calculation performed by the returned function.
Calculation names are case-insensitive.",
                    Documentation.AllowedValues = availableCalculations
                ],
            optionsParameter = type nullable record
                meta [
                    Documentation.FieldCaption = "Options",
                    Documentation.FieldDescription =
                        "Optional record controlling how the calculation is performed.
Supported fields can be supplied in any order:
• Rounding — ROUND (default), ROUNDUP, or ROUNDDOWN. Names are case-insensitive.
• Digits — whole number of decimal digits. Negative values round left of the decimal point.
• Step — non-zero multiple for MROUND-style rounding. The absolute magnitude is used.
• IgnoreNull — false by default. Set true to remove null arguments before calculation.
• Culture — culture used by Number.From when converting supplied values. The default is invariant culture.

Step takes precedence when both Step and Digits are supplied.
Rounding has no effect unless either Digits or Step is supplied."
                ],
            functionType = type function (
                calculation as calculationParameter,
                optional options as optionsParameter
            ) as function
                meta [
                    Documentation.Name = "fxExcelStyleCalc",
                    Documentation.LongDescription =
                        "Creates a configured calculation function using Excel-style calculations and rounding.
The required 'calculation' argument specifies what to do. The optional 'options' record specifies how to do it.
Arguments supplied to the returned function are converted through Number.From before calculation. Nulls propagate by default or can be removed with IgnoreNull = true.
ROUND uses midpoint rounding away from zero. Step provides MROUND-style rounding to a multiple and can also be combined with ROUNDUP or ROUNDDOWN.
The function deliberately retains Power Query M behavior where appropriate. Native calculation and conversion errors are not replaced, and special numeric values such as NaN and infinity remain numeric values rather than being converted to Excel worksheet errors.
This is Excel-style behavior, not a strict Excel calculation emulator.",
                    Documentation.Category = "Number",
                    Documentation.Version = "1.00: Initial Version",
                    Documentation.Author = "Melissa de Korte",
                    Documentation.Examples = {
                        [
                            Description = "Sum several values using the default options.",
                            Code = "let calc = fxExcelStyleCalc(""SUM"") in calc(10, 20, 30)",
                            Result = "60"
                        ],
                        [
                            Description = "Calculate an average and round the result to two decimal places.",
                            Code = "let calc = fxExcelStyleCalc(""AVERAGE"", [Digits = 2]) in calc(10.125, 20.125, 30.125)",
                            Result = "20.13"
                        ],
                        [
                            Description = "Round a calculated result to the nearest multiple of 5.",
                            Code = "let calc = fxExcelStyleCalc(""SUM"", [Step = 5]) in calc(11, 6)",
                            Result = "15"
                        ]
                    }
                ],
            addFunctionType = Value.ReplaceType(
                fxExcelStyleCalc,
                functionType
            )
        in
            addFunctionType
in
    fxWithDocumentation
```

---

## Function shape

The outer function separates a required calculation from its optional behavior:

```m
(calculation as text, optional options as nullable record) as function
```

Conceptually:

```m
fxExcelStyleCalc(
    what to do,
    how to do it
)
```

The first argument selects the calculation.  
The second argument is an optional record containing settings that modify how that calculation behaves.

Assuming the query is named `fxExcelStyleCalc`:

```m
calc = fxExcelStyleCalc(
    "AVERAGE",
    [Digits = 2]
)
```

`calc` is now a configured function that calculates an average, rounded to two decimal places:

```m
calc(10.125, 20.125, 30.125)
```

Result:

```text
20.13
```

`calc` is a **closure**: a function that retains the selected calculation, rounding method, culture, null policy, and rounding precision from the arguments provided when it was created.

The calculation is required explicitly telling you its purpose:

```m
fxExcelStyleCalc("SUM")
fxExcelStyleCalc("AVERAGE")
fxExcelStyleCalc("DIVIDE")
```

This keeps the purpose clear for every configuration that is initialized.

---

## Configuration options

The first argument, `calculation`, is required and case-insensitive. It accepts one of the calculation names listed in the next section.

The second argument, `options`, is optional. All fields inside the options record are optional, and `Rounding` is case-insensitive.

| Option | Default | Accepted value | Purpose |
|---|---|---|---|
| `Rounding` | `"ROUND"` | `"ROUND"`, `"ROUNDUP"`, `"ROUNDDOWN"` | Selects how a result is rounded when `Digits` or `Step` is supplied. |
| `Digits` | `null` | Whole number, including negative values | Rounds to that number of decimal digits. Negative values round to positions left of the decimal point. |
| `Step` | `null` | Non-zero numeric value | Rounds to a multiple. Its magnitude is used, so `-5` and `5` both mean a step of `5`. |
| `IgnoreNull` | `false` | `true` or `false` | Controls whether null arguments are removed or cause a null result. |
| `Culture` | `""` | Power Query culture name such as `"en-US"` or `"de-DE"` | Controls how arguments are interpreted by `Number.From`. |

The empty culture string is the invariant culture, giving the function predictable parsing unless another culture is explicitly requested.

If both `Step` and `Digits` are supplied, **`Step` takes precedence**.  
If neither is supplied, no rounding is performed. Merely setting `Rounding = "ROUNDUP"` does not cause the result to be rounded.

Unknown fields in the options record are ignored because settings are retrieved with `Record.FieldOrDefault`.

---

## Supported calculations

| Calculation | Arguments | Implementation / behavior |
|---|---:|---|
| `SUM` | 1…n | Sum of all arguments |
| `AVERAGE` | 1…n | Arithmetic mean |
| `PRODUCT` | 1…n | Product of all arguments |
| `MIN` | 1…n | Minimum |
| `MAX` | 1…n | Maximum |
| `MEDIAN` | 1…n | Median |
| `SUBTRACT` | 1…n | Sequential subtraction |
| `DIVIDE` | 1…n | Sequential division |
| `POWER` | 2 | Raises the first value to the second |
| `SQRT` | 1 | Square root |
| `EXP` | 1 | Raises *e* to the supplied value |
| `FACT` | 1 | Factorial after truncating the value toward zero |
| `ABS` | 1 | Absolute value |
| `SIGN` | 1 | Returns `-1`, `0`, or `1` |
| `LN` | 1 | Natural logarithm |
| `LOG10` | 1 | Base-10 logarithm |
| `LOG` | 1–2 | Logarithm; defaults to base 10 |
| `MOD` | 2 | Excel-style modulo result |

### Excel-style adjustments

`LOG` deliberately supplies `10` when its optional base is absent because native M `Number.Log` defaults to *e*, whereas Excel `LOG` defaults to base 10.

`MOD` deliberately uses:

```m
n - d * Number.RoundDown(n / d)
```

This follows Excel's `MOD` definition and gives the remainder the sign of the divisor. Runtime testing showed that native `Number.Mod` differs for negative operands.

`FACT` first truncates toward zero:

```m
Number.Factorial(Number.RoundTowardZero(_))
```

This matches Excel-style treatment of a fractional factorial argument.

---

## Why `Function.From` is central to the design

The calculation engine works internally with a list, but callers do not have to construct one.

For example:

```m
sum = fxExcelStyleCalc("SUM"),

result = sum(10, 20, 30, 40)
```

The caller supplies separate arguments. `Function.From` packages the supplied arguments into the list consumed by the internal calculation.

That lets native list functions become convenient multi-argument functions:

```m
SUM = f(List.Sum),
AVERAGE = f(List.Average),
PRODUCT = f(List.Product)
```

Conceptually:

```text
sum(10, 20, 30)
        ↓
{10, 20, 30}
        ↓
List.Sum
        ↓
60
```

The open-ended extra-argument behavior used here has been observed in the target M environment. Microsoft's documentation describes the argument-to-list mechanism but does not explicitly document unlimited arguments beyond the parameters shown by the supplied function type.

---

## Sequential calculations

Subtraction and division cannot safely be simplified algebraically when floating-point numbers are involved.

For example:

```text
1E16 - 1E16 - 1
```

must remain:

```text
(1E16 - 1E16) - 1
```

rather than:

```text
1E16 - (1E16 + 1)
```

Those expressions are mathematically equivalent over exact real numbers, but not necessarily under finite floating-point precision.

For that reason, `SUBTRACT` and `DIVIDE` use:

```m
fold = (fn as function) as function =>
    f(each List.Accumulate(List.Skip(_), _{0}, fn))
```

with:

```m
SUBTRACT = fold(Value.Subtract),
DIVIDE = fold(Value.Divide)
```

`List.Accumulate` carries each intermediate result forward to the next operation, preserving sequential evaluation.

---

## Input conversion

Every supplied argument is converted before calculation using:

```m
Number.From(_, culture)
```

`Number.From` accepts numbers as well as convertible values such as numeric text and logical values. Unsupported values raise their natural M conversion error.

For example:

```m
sum = fxExcelStyleCalc("SUM"),

result = sum("10", "20.5", 4.5)
```

returns:

```text
35
```

with the default invariant culture.

A culture can be supplied when the text uses another convention:

```m
sumDE = fxExcelStyleCalc(
    "SUM",
    [Culture = "de-DE"]
),

result = sumDE("1,5", "2,5")
```

returns the numeric value:

```text
4
```

`Culture` controls **conversion**, not the data type of the returned result. A successful result remains a Power Query `number`.

---

## Null handling

Null behavior is explicit rather than inherited from the selected calculation.

By default:

```m
IgnoreNull = false
```

so:

```m
sum = fxExcelStyleCalc("SUM"),

result = sum(10, null, 20)
```

returns:

```text
null
```

No calculation or rounding is performed.

With:

```m
sum = fxExcelStyleCalc(
    "SUM",
    [IgnoreNull = true]
)
```

the same arguments:

```m
sum(10, null, 20)
```

return:

```text
30
```

If all supplied values are removed because they are null, the function returns `null`.

This is a deliberate choice rather than an attempt to reproduce every Excel worksheet function's treatment of blank cells.

---

## Excel-style rounding

Three rounding modes are available:

| Setting | Behavior |
|---|---|
| `ROUND` | Nearest value; midpoint ties are rounded away from zero |
| `ROUNDUP` | Always away from zero |
| `ROUNDDOWN` | Always toward zero |

For `ROUND`, the function explicitly supplies:

```m
RoundingMode.AwayFromZero
```

because native `Number.Round` otherwise uses ties-to-even. This gives the familiar Excel `ROUND` midpoint behavior.

### Rounding by digits

```m
calc = fxExcelStyleCalc(
    "AVERAGE",
    [
        Rounding = "ROUND",
        Digits = 2
    ]
),

result = calc(10.125, 20.125, 30.125)
```

Result:

```text
20.13
```

Positive, zero, and negative digits are supported:

```m
[Digits = 2]     // hundredths
[Digits = 0]     // integers
[Digits = -1]    // tens
[Digits = -2]    // hundreds
```

`Digits` must resolve through `Number.From` to a whole number. A fractional value raises:

```text
Digits must be a whole number.
```

---

## Rounding by Step — MROUND-style multiple rounding

`Step` provides **MROUND-style rounding to a multiple**.

Excel `MROUND(number, multiple)` rounds a number to the nearest requested multiple. This function uses the same core idea:

```text
value ÷ step
      ↓
round to an integer
      ↓
× step
```

For example:

```m
calc = fxExcelStyleCalc(
    "SUM",
    [Step = 5]
),

result = calc(11, 6)
```

The calculation produces `17`, which is rounded to the nearest multiple of `5`:

```text
15
```

Other examples:

```m
[Step = 0.05]
[Step = 0.25]
[Step = 10]
[Step = 100]
```

### Relationship to Excel `MROUND`

When:

```m
Rounding = "ROUND"
```

`Step` is best understood as an **MROUND-style mode**: round the result to the nearest multiple.

There are two deliberate differences from Excel's exact `MROUND` contract:

1. **Step is treated as a positive magnitude.**  
   The function applies `Number.Abs`, so:

   ```m
   [Step = -10]
   ```

   behaves like:

   ```m
   [Step = 10]
   ```

   Excel `MROUND` requires `number` and `multiple` to have the same sign and returns `#NUM!` when their signs differ.

2. **The same multiple-rounding mechanism is extended to `ROUNDUP` and `ROUNDDOWN`.**  
   This provides directed multiple rounding:

   ```m
   [
       Rounding = "ROUNDUP",
       Step = 10
   ]
   ```

   rounds away from zero to a multiple of `10`, while:

   ```m
   [
       Rounding = "ROUNDDOWN",
       Step = 10
   ]
   ```

   rounds toward zero to a multiple of `10`.

   These are useful extensions of the Step mechanism; they are not claims that Excel's `MROUND` itself has `ROUNDUP` or `ROUNDDOWN` modes.

Step scaling uses `Value.Divide` and `Value.Multiply` with `Precision.Decimal` to reduce avoidable floating-point artifacts introduced by divide-and-multiply rounding.

A zero Step raises:

```text
Step cannot be zero.
```

---

## Configuration examples

### Default: SUM, no rounding

```m
sum = fxExcelStyleCalc("SUM")
```

```m
sum(10, 20, 30)
// 60
```

### Average rounded to two decimal places

```m
average2 = fxExcelStyleCalc(
    "AVERAGE",
    [Digits = 2]
)
```

### MROUND-style rounding to the nearest multiple of 5

```m
nearest5 = fxExcelStyleCalc(
    "SUM",
    [Step = 5]
)
```

### Round away from zero to a multiple of 10

```m
upTo10 = fxExcelStyleCalc(
    "SUM",
    [
        Rounding = "ROUNDUP",
        Step = 10
    ]
)
```

### Ignore nulls and parse German-formatted numeric text

```m
germanAverage = fxExcelStyleCalc(
    "AVERAGE",
    [
        IgnoreNull = true,
        Culture = "de-DE",
        Digits = 2
    ]
)
```

### Excel-style LOG with an optional base

```m
log = fxExcelStyleCalc("LOG")

base10 = log(100)
base2 = log(8, 2)
```

---

## Error and evaluation behavior

The function intentionally does not protect calculations, or replace errors.

If input conversion fails, the native conversion error propagates. If the selected calculation fails, that error propagates.

This keeps errors diagnostic: the caller receives the error raised by the operation that failed.

`NaN` and positive or negative infinity are different. They are numeric values in M rather than errors, so native operations that return one of those values continue through the pipeline.

---

## Design decisions

### 1. The required action is separated from optional behavior

The first argument answers **what to do** by selecting the calculation. The optional record answers **how to do it** by supplying rounding, null, and culture settings.

This keeps the primary intent visible:

```m
fxExcelStyleCalc(
    "AVERAGE",
    [Digits = 2]
)
```

rather than burying the calculation name in a record. The returned closure then performs that configured operation repeatedly.

### 2. Native M functions are preferred

Where an M function already provides the required behavior, it is used directly:

```m
POWER = Number.Power
SQRT = Number.Sqrt
EXP = Number.Exp
ABS = Number.Abs
SIGN = Number.Sign
```

Additional code is used only where the Excel-style contract requires a meaningful adjustment.

### 3. Variadic calculations use `Function.From`

Functions such as `SUM`, `AVERAGE`, and `PRODUCT` naturally consume lists in M. `Function.From` adapts them so callers can supply separate arguments.

### 4. Fixed-arity calculations use `Function.Invoke`

After conversion and null handling, the selected calculation is invoked from the argument list:

```m
Function.Invoke(calc, values)
```

This lets the same dispatcher handle both variadic wrappers and fixed-arity native number functions.

### 5. Sequential operations preserve evaluation order

`SUBTRACT` and `DIVIDE` use `List.Accumulate` with `Value.Subtract` and `Value.Divide` so that floating-point behavior follows sequential evaluation rather than an algebraically rearranged expression.

### 6. Null handling is explicit

`IgnoreNull` makes null behavior a deliberate configuration choice rather than depending on explicit null handling semantics.

### 7. Natural errors are retained

The function avoids `try` in the calculation pipeline so that native M errors are preserved and evaluation stops when an operation fails.

### 8. Rounding happens after calculation

Inputs are converted and calculated first. Rounding is applied to the final calculated value only.

### 9. `Step` is MROUND-style, not a strict MROUND clone

`ROUND + Step` follows the familiar Excel idea of rounding to a multiple, while the function deliberately treats the step as a magnitude and extends the mechanism to `ROUNDUP` and `ROUNDDOWN`.

---

## Excel-style, not strict Excel equivalence

The intended description is:

> **Excel-style calculations and rounding**

The function uses familiar Excel calculation names and deliberately aligns important rounding semantics, while still behaving as a Power Query M function.

Notable distinctions include:

- every supplied argument is normalized through `Number.From`;
- null handling is explicitly configurable;
- native M errors are preserved;
- special numeric values such as `NaN` and infinity remain M numeric values;
- `Step` treats a negative multiple as a positive magnitude;
- `ROUNDUP` and `ROUNDDOWN` can also be applied to Step multiples;
- some M numeric edge cases may differ from Excel worksheet errors.

This makes the function familiar to Excel users without turning it into a full Excel calculation emulator.

---
