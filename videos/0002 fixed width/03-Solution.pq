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
    headers = List.ReplaceRange( List.Transform(splitter(header), Text.Trim), 6, 2, {"Last Cost Totals", "Standard Costs Totals"}),
    
    /* DATA */
    Source = RAW_DATA,
    filter = Table.SelectRows(Source, each Text.Length([Column1]) = rowLen and Text.StartsWith([Column1], "PC")),
    split = Table.SplitColumn(filter, "Column1", splitter, headers),
    chType = Table.TransformColumnTypes(split,{{"Quantity", Int64.Type}, {"Last Cost", Currency.Type}, {"Standard Cost", Currency.Type}, {"Difference", Currency.Type}, {"Last Cost Totals", Currency.Type}, {"Standard Costs Totals", Currency.Type}}, "en-US")
in
    chType
