let
    Source = Table.FromValue({ 
    (_) => _
,     each _ 
})
in
    Source
