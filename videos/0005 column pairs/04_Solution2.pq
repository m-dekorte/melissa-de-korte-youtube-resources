/* Solution2 uses: RAW_DATA2 */
/* Works when all Hours columns appear first, followed by all Cost columns */
/* Only colSets changes relative to Solution1 — everything else is the same */

let
    /* PREP */
    otCols = List.Select(Table.ColumnNames(Source), each Text.StartsWith(_, "OT ")),
    colSets = List.Zip(List.Split(otCols, List.Count(otCols) / 2)),

    /* DATA */
    Source = RAW_DATA2,
    pairs = Table.AddColumn(Source, "_rows", each
        List.Transform(colSets, (c) =>
            Record.FromList(
                {"OT @" & Text.AfterDelimiter(List.First(c), "@")} &
                List.Transform(c, (v) => Record.Field(_, v)),
                {"Rate Type", "Hours", "Cost"}
            )
        ), type {[Rate Type = text, Hours = number, Cost = number]}
    ),
    removeCols = Table.RemoveColumns(pairs, otCols),
    expandLists = Table.ExpandListColumn(removeCols, "_rows"),
    expandRecords = Table.ExpandRecordColumn(expandLists, "_rows", {"Rate Type", "Hours", "Cost"})
in
    expandRecords
