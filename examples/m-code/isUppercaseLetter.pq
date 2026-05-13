let
    isUppercaseLetter = (char as nullable text) as logical => 
        if Text.Length(char ?? "") > 1 then error "#(lf,tab)Expected a single character, but the value provided contains " & Text.From(Text.Length(char)) & " characters."
        else Text.Lower(char) <> char,
    isUppercaseLetterType = type function (
        char as (
            type nullable text
                meta [
                    Documentation.FieldCaption = "char",
                    Documentation.FieldDescription = "A single character to test for uppercase. For example: ""A""."
                ]
        )
    ) as logical
        meta [
            Documentation.Name = "isUppercaseLetter",
            Documentation.LongDescription = "Returns true when the input character changes after converting it to lowercase, which indicates it is an uppercase letter.",
            Documentation.Category = "Text",
            Documentation.Version = "1.00: initial version",
            Documentation.Author = " Melissa de Korte",
            Documentation.Examples = {
                [
                    Description = "Uppercase letter",
                    Code        = "isUppercaseLetter(""A"")",
                    Result      = "true"
                ],
                [
                    Description = "Lowercase letter",
                    Code        = "isUppercaseLetter(""a"")",
                    Result      = "false"
                ],
                [
                    Description = "Digit",
                    Code        = "isUppercaseLetter(""1"")",
                    Result      = "false"
                ]
            }
        ]
in
    Value.ReplaceType(isUppercaseLetter, isUppercaseLetterType)
