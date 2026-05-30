# Parsing Fixed-Width Data in Power Query

A step-by-step guide to building a dynamic fixed-width file parser that detects column boundaries automatically instead of hard-coding positions.

**Video:** [Watch on YouTube](https://www.youtube.com/@melissa_de_korte)  
**LinkedIn:** [Connect with Melissa](https://www.linkedin.com/in/melissa-de-korte-64585884/)

---

## The Problem

Most people hard-code column positions when parsing fixed-width files. But when you get different reports from the same source system, you end up building solutions from scratch each time. 
Instead, create building blocks, pieces of code that can easily be reused and adapted to suit your needs.

---

## The Solution: Recognition + Detection

When a text file has no obvious separator, switch to a monospaced font. If columns suddenly line up perfectly, you're dealing with fixed-width data.

Our sample data had a divider row—the one with dashes—it shows exactly where every column begins. If we can determine those boundaries automatically, we can split all rows using those positions. Should the file format change, the divider changes too. Our parser detects the new positions and adapts automatically.

---

## Setup

### 1. Create the FileLocation Parameter Query

In the Power Query Editor, create a new blank query named `FileLocation`:

```m
"C:\Downloads\sampleFixedWidth.txt" meta [
    IsParameterQuery=true, 
    List={"C:\Downloads\sampleFixedWidth.txt"}, 
    DefaultValue="C:\Downloads\sampleFixedWidth.txt", 
    Type="Text", 
    IsParameterQueryRequired=true
]
```

Then edit the parameter and set it to the actual path of the sample fixed-width file on your machine.

### 2. Create the RAW_DATA Query

Create a new blank query named `RAW_DATA`:

```m
let
    Source = Table.FromColumns({Lines.FromBinary(File.Contents(FileLocation), null, null, 1252)})
in
    Source
```

This imports the sample text file as a single column. Each row is one line of the file. Everything that follows builds on this raw data.

### 3. Create the Solution Query

Create a new blank query named `Solution` (or whatever you prefer). This is where you'll build the parser step by step. The following sections walk through each piece.

---

## Building the Query: Step by Step

All the logic below goes into your Solution query inside the `let...in` statement.

### PREP SECTION: Building Tools

#### Step 1: Reference Your Raw Data

```m
Source = RAW_DATA
```

This references a query, that connects to the sample file and returns a single-column table.

Place the following steps, above the Source step in your query.

---

#### Step 2: Find the Header Row

```m
header = Table.Skip(Source, each not Text.Contains([Column1], "Product")){0}[Column1],
```

This searches for the row containing "Product" and extracts that text. `Table.Skip` keeps skipping rows as long as the condition is true, then stops. `{0}[Column1]` means: take the first row, and extract the text from Column1.

**Why not hard-code row numbers?** This pattern is designed to be dynamic and reusable across different solutions. Instead of relying on fixed row numbers, it searches for specific text to identify the correct position in the data.

When applying this pattern to another file or scenario, you usually only need to update the text to search for, such as replacing "Product", and, if needed, the column to search in, such as `[Column1]`.

Searching for text makes the solution more flexible and easier to adapt when the layout changes.

---

#### Step 3: Find the Divider Row

```m
divider = Table.Skip(Source, each not Text.Contains([Column1], "-- ")){0}[Column1],
```

Same pattern. This row (containing dashes) shows exactly where columns begin and end. It's your key to automatic position detection.

---

#### Step 4: Measure Row Length

```m
rowLen = Text.Length(divider),
```

In fixed-width format, all valid data rows are the same length. This becomes a filter rule later: only include rows that match this length.

---

#### Step 5: Detect Column Positions

This is the core logic. It turns the divider into a list of positions.

```m
findStarts = (s) => List.Select(
    {0} & List.Transform(Text.PositionOf(s, " ", Occurrence.All), each _+1),
    each Text.At(s, _) <> " "
),
```

**How it works (inside out):**

1. `Text.PositionOf(s, " ", Occurrence.All)` — Returns a list with the position of every space in the string
2. `List.Transform(..., each _+1)` — Add 1 to each position (columns start after spaces)
3. `{0} &` — Prepend 0 (the first column always starts at position 0, it has no preceding dashes and space)
4. `List.Select(..., each Text.At(s, _) <> " ")` — Filter out any position that points to a space character

**Result:** A list of valid column start positions, like `{0, 10, 20, 30, 40, 45, 55}`.

---

#### Step 6: Apply the Function to Get Positions

```m
positions = findStarts(divider),
```

Detect the positions from your divider and store them.

---

#### Step 7: Create a Reusable Splitter

```m
splitter = Splitter.SplitTextByPositions(positions),
```

This creates a splitter function that will cut text at your detected positions. You'll use it on both the header and every data row.

---

#### Step 8: Extract and Clean Headers

```m
headers = List.ReplaceRange(
    List.Transform(splitter(header), Text.Trim),
    6, 2,
    {"Last Cost Totals", "Standard Costs Totals"}
),
```

1. `splitter(header)` — Split the header using your positions
2. `List.Transform(..., Text.Trim)` — Remove extra spaces from each column name
3. `List.ReplaceRange(..., 6, 2, {...})` — Replace the values at (zero-based) positions 6 and 7 with meaningful names

In your sample file, the header row spans two physical lines, one was extracted. When split, some columns end up with duplicate names. This line renames them to match the actual column contents.

---

### DATA SECTION: Applying the Tools

Place the following steps below the Source step.

After a step is incorporated, return its name below the in-clause to review the outcome.

#### Step 9: Filter to Valid Data Rows

```m
    filter = Table.SelectRows(Source, each Text.Length([Column1]) = rowLen and Text.StartsWith([Column1], "PC"))
in
    filter
```

Two conditions:

1. **Row length matches:** `Text.Length([Column1]) = rowLen` — Only include rows that are exactly the same length as the divider
2. **Starts with product code:** `Text.StartsWith([Column1], "PC")` — Only include rows starting with "PC" (the product code prefix)

This filters out headers, dividers, blank lines, and malformed rows.

---

#### Step 10: Split by Position

Add a comma at the end of the previous step, fiter. 

```m
    split = Table.SplitColumn(filter, "Column1", splitter, headers)
in
    split
```

Uses your splitter to cut each filtered row into columns, with names from your cleaned headers list.

---

#### Step 11: Apply Column Types

```m
    chType = Table.TransformColumnTypes(split,
        {
            {"Quantity", Int64.Type}, 
            {"Last Cost", Currency.Type}, 
            {"Standard Cost", Currency.Type}, 
            {"Difference", Currency.Type}, 
            {"Last Cost Totals", Currency.Type}, 
            {"Standard Costs Totals", Currency.Type}
        }, 
        "en-US"
    )
in
    chType
```

Explicitly sets the data type for each column:
- **Quantity** → Whole number
- **Cost columns** → Currency
- Culture is `"en-US"` so Power Query interprets values correctly

All values came in as text. This step tells Power Query: cost columns are currency, and quantity is an integer.

Once the optional Culture parameter is set, setting a column type will generate a new Changed Type step. 

---

## Complete Query (Copy-Paste Reference)

```m
let
    /* PREP */
    header = Table.Skip(Source, each not Text.Contains([Column1], "Product")){0}[Column1],
    divider = Table.Skip(Source, each not Text.Contains([Column1], "-- ")){0}[Column1],
    rowLen = Text.Length(divider),
    findStarts = (s) => List.Select(
        {0} & List.Transform(Text.PositionOf(s, " ", Occurrence.All), each _+1),
        each Text.At(s, _) <> " "
    ),
    positions = findStarts(divider),
    splitter = Splitter.SplitTextByPositions(positions),
    headers = List.ReplaceRange(
        List.Transform(splitter(header), Text.Trim),
        6, 2,
        {"Last Cost Totals", "Standard Costs Totals"}
    ),
    
    /* DATA */
    Source = RAW_DATA,
    filter = Table.SelectRows(Source, each Text.Length([Column1]) = rowLen and Text.StartsWith([Column1], "PC")),
    split = Table.SplitColumn(filter, "Column1", splitter, headers),
    chType = Table.TransformColumnTypes(split,
        {
            {"Quantity", Int64.Type}, 
            {"Last Cost", Currency.Type}, 
            {"Standard Cost", Currency.Type}, 
            {"Difference", Currency.Type}, 
            {"Last Cost Totals", Currency.Type}, 
            {"Standard Costs Totals", Currency.Type}
        }, 
        "en-US"
    )
in
    chType
```

---

---

## Testing Before Production

1. **Import a new export** — Test with a new export to make sure the solution is stable
2. **Preview filtered rows** — Check what you're removing by inverting the logical expression
3. **Verify totals** — Sum numeric columns to confirm value conversion is correct and all data is present

---

## Key Takeaway

Fixed-width files follow a pattern: look for somthing such as a divider, to detect boundaries automatically. Once you found the pattern, the parsing becomes straightforward instead of manual. When the format changes, your query adapts to the new structure.

Think about a modular appraoch, that enables you to quickly deploy mini-solutions (steps/ building blocks) across different projects. 

---

## Resources

- **Sample File:** `00-sampleFixedWidth.txt`
- **All Three Queries:** `FileLocation`, `RAW_DATA`, `Solution`
- **Video Tutorial:** [YouTube](https://www.youtube.com/@melissa_de_korte)
- **Full Step-by-Step Guide:** See sections above
- **Questions?** [Connect on LinkedIn](https://www.linkedin.com/in/melissa-de-korte-64585884/)

Happy parsing!
