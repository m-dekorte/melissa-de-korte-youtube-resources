let
    fxPromoteMultiRowHeaders =
    let 
        fnPromoteMultiRowHeaders  =(t as table, startRow as number, rowCount as number) as table =>
        let
            offset = startRow -1,
            headers = List.Transform(Table.ToColumns(Table.Range(t, offset, rowCount)), each
                let x  = List.RemoveMatchingItems(_, {"", null}) in
                if List.IsEmpty(x) then null else Text.Combine(List.Transform(x, Text.From), " ")
            ),
            transform = Table.PromoteHeaders(
                Table.InsertRows(Table.Skip(t, offset+rowCount), 0, 
                    {Record.FromList(headers, Table.ColumnNames(t))}
                )
            )
        in
            transform,
        TableParameter = type table meta[
                Documentation.FieldCaption = "t",
                Documentation.FieldDescription = "The source table containing the rows to combine and promote as column headers."
            ],
            StartRowParameter = type number meta [
                Documentation.FieldCaption = "startRow",
                Documentation.FieldDescription = "The 1-based position of the first header row. Rows above this position are discarded from the result."
            ],
            RowCountParameter = type number meta [
                Documentation.FieldCaption = "rowCount",
                Documentation.FieldDescription = "The number of consecutive rows, beginning at startRow, to combine into column headers."
            ],
            AddParameterTypes = type function (
                t as TableParameter,
                startRow as StartRowParameter,
                rowCount as RowCountParameter
            ) as table meta [
                Documentation.Name = "fxPromoteMultiRowHeaders",
                Documentation.Description = "Combines values from multiple consecutive rows to headers.",
                Documentation.LongDescription = "Combines values from multiple consecutive rows into one header for each column. Nulls and empty text values are removed before the remaining values are converted to text and joined with a single space. Rows preceding and comprising the header block are removed, and the combined values are promoted as column names. See: https://github.com/m-dekorte/melissa-de-korte-youtube-resources/tree/main/videos/0006%20multi%20row%20headers",
                Documentation.Category = "Table",
                Documentation.Version = "1.00: Initial Version",
                Documentation.Author = "Melissa de Korte"
            ]
        in
            Value.ReplaceType(fnPromoteMultiRowHeaders, AddParameterTypes)
in
    fxPromoteMultiRowHeaders
