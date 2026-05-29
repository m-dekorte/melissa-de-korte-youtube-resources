let
    Source = Table.FromColumns({Lines.FromBinary(File.Contents(FileLocation), null, null, 1252)})
in
    Source
