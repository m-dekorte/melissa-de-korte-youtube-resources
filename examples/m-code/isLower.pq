let
    /*
        Created by: Melissa de Korte
        Subscribe: https://www.youtube.com/@melissa_de_korte
        Resource: https://github.com/m-dekorte/melissa-de-korte-youtube-resources/blob/main/examples/m-code/isLower.pq
    */
    isLower = (char as nullable text) as logical =>
        if Text.Length(char ?? "") > 1 then error "#(lf,tab)Expected a single character, but the value provided contains " & Text.From(Text.Length(char)) & " characters."
        else Text.Upper(char) <> char,
    isLowerType = type function (
        char as (
            type nullable text
                meta [
                    Documentation.FieldCaption = "char",
                    Documentation.FieldDescription = "A single character to test for lowercase. For example: ""a""."
                ]
        )
    ) as logical
        meta [
            Documentation.Name = "isLower",
            Documentation.LongDescription = "Returns true when the input character changes after converting it to uppercase, which indicates it is a lowercase letter.",
            Documentation.Category = "Text",
            Documentation.Version = "1.00: initial version",
            Documentation.Author = " Melissa de Korte",
            Documentation.Examples = {
                [
                    Description = "Lowercase letter",
                    Code        = "isLower(""a"")",
                    Result      = "true"
                ],
                [
                    Description = "Uppercase letter",
                    Code        = "isLower(""A"")",
                    Result      = "false"
                ],
                [
                    Description = "Digit",
                    Code        = "isLower(""1"")",
                    Result      = "false"
                ]
            }
        ]
in
    Value.ReplaceType(isLower, isLowerType)
