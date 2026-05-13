let
    isLowercaseLetter = (char as nullable text) as logical =>
        if Text.Length(char ?? "") > 1 then error "#(lf,tab)Expected a single character, but the value provided contains " & Text.From(Text.Length(char)) & " characters."
        else Text.Upper(char) <> char,
    isLowercaseLetterType = type function (
        char as (
            type nullable text
                meta [
                    Documentation.FieldCaption = "char",
                    Documentation.FieldDescription = "A single character to test for lowercase. For example: ""a""."
                ]
        )
    ) as logical
        meta [
            Documentation.Name = "isLowercaseLetter",
            Documentation.LongDescription = "Returns true when the input character changes after converting it to uppercase, which indicates it is a lowercase letter.",
            Documentation.Category = "Text",
            Documentation.Examples = {
                [
                    Description = "Lowercase letter",
                    Code        = "isLowercaseLetter(""a"")",
                    Result      = "true"
                ],
                [
                    Description = "Uppercase letter",
                    Code        = "isLowercaseLetter(""A"")",
                    Result      = "false"
                ],
                [
                    Description = "Digit",
                    Code        = "isLowercaseLetter(""1"")",
                    Result      = "false"
                ]
            }
        ]
in
    Value.ReplaceType(isLowercaseLetter, isLowercaseLetterType)
