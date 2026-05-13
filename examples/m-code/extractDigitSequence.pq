let
    extractDigitSequence = (txt as nullable text, len as number, optional Occurrence as number) =>
        let
            input  = txt ?? "",
            parts  = Text.SplitAny(input, Text.Remove(input, {"0" .. "9"})),
            digits = List.Select(parts, each Text.Length(_) = len)
        in
            {List.First(digits), List.Last(digits), digits}{(Occurrence ?? 0)},
    allowedLen = {1 .. 255},
    extractDigitSequenceType = type function (
        txt as (
            type nullable text
                meta [
                    Documentation.FieldCaption = "txt",
                    Documentation.FieldDescription = "A text string to scan for n-digit sequences (e.g. ""Sales20250109T12_45.csv""). "
                ]
        ),
        len as (
            type number
                meta [
                    Documentation.FieldCaption = "len",
                    Documentation.FieldDescription = "The digit sequence length to scan for in a string (e.g. 8). ",
                    Documentation.AllowedValues = allowedLen
                ]
        ),
        optional Occurrence as (
            type number
                meta [
                    Documentation.FieldCaption = "Occurrence",
                    Documentation.FieldDescription = "Controls which value to return: 
0 – Occurrence.First, the first n-digit match (default)
1 – Occurrence.Last, the last n-digit match
2 – Occurrence.All, all n-digit matches as a list",
                    Documentation.AllowedValues = {0, 1, 2}
                ]
        )
    ) as any
        meta [
            Documentation.Name = "extractDigitSequence",
            Documentation.LongDescription = "Returns only n digit sequences from the input string. Defaults to occurrence first.",
            Documentation.Category = "Text",
            Documentation.Version = "1.10: Added a len parameter",
            Documentation.Author = " Melissa de Korte",
            Documentation.Examples = {
                [
                    Description = "Default (occurrence first)",
                    Code        = "extractDigitSequence(""Sales20250109T12_45.csv"", 8)",
                    Result      = "20250109"
                ],
                [
                    Description = "Occurance all",
                    Code = "extractDigitSequence(""US 01032025 GB 03012025 EU 05012025.csv"", 8, 2)",
                    Result = "{ ""01032025"", ""03012025"", ""05012025"" }"
                ]
            }
        ]
in
    Value.ReplaceType(extractDigitSequence, extractDigitSequenceType)
