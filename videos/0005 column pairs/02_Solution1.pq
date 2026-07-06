/* Solution1 uses: RAW_DATA */
/* Works when OT columns are paired repeadedly: OT Hours @x followed by OT Cost @x */

let
    /* PREP */
    otCols = List.Select(Table.ColumnNames(Source), each Text.StartsWith(_, "OT ")),
    colSets = List.Split(otCols, 2),

    /* DATA */
    Source = RAW_DATA,
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
