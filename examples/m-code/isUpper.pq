let
    /*
        Created by: Melissa de Korte
        Subscribe: https://www.youtube.com/@melissa_de_korte
        Resource: https://github.com/m-dekorte/melissa-de-korte-youtube-resources/blob/main/examples/m-code/isUpper.pq
    */
    isUpper = (char as nullable text) as logical => 
        if Text.Length(char ?? "") > 1 then error "#(lf,tab)Expected a single character, but the value provided contains " & Text.From(Text.Length(char)) & " characters."
        else Text.Lower(char) <> char,
    isUpperType = type function (
        char as (
            type nullable text
                meta [
                    Documentation.FieldCaption = "char",
                    Documentation.FieldDescription = "A single character to test for uppercase. For example: ""A""."
                ]
        )
    ) as logical
        meta [
            Documentation.Name = "isUpper",
            Documentation.LongDescription = "Returns true when the input character changes after converting it to lowercase, which indicates it is an uppercase letter.",
            Documentation.Category = "Text",
            Documentation.Version = "1.00: initial version",
            Documentation.Author = " Melissa de Korte",
            Documentation.Examples = {
                [
                    Description = "Uppercase letter",
                    Code        = "isUpper(""A"")",
                    Result      = "true"
                ],
                [
                    Description = "Lowercase letter",
                    Code        = "isUpper(""a"")",
                    Result      = "false"
                ],
                [
                    Description = "Digit",
                    Code        = "isUpper(""1"")",
                    Result      = "false"
                ]
            }
        ]
in
    Value.ReplaceType(isUpper, isUpperType)
