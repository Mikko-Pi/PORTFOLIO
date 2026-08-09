let
    Source = Folder.Files("C:\Users\mikko\OneDrive\Tiliotteet"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"Id", type text}, {"Name", type text}, {"Kind", type text}, {"Data", type any}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Kind] = "Page")),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows",".pdf","",Replacer.ReplaceText,{"Source.Name"}),
    #"Expanded Data" = Table.ExpandTableColumn(#"Replaced Value", "Data", {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12"}, {"Data.Column1", "Data.Column2", "Data.Column3", "Data.Column4", "Data.Column5", "Data.Column6", "Data.Column7", "Data.Column8", "Data.Column9", "Data.Column10", "Data.Column11", "Data.Column12"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Data",{"Id", "Name", "Kind", "Data.Column6", "Data.Column7", "Data.Column8", "Data.Column9", "Data.Column10", "Data.Column11", "Data.Column12", "Data.Column2", "Source.Name"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Removed Columns",{{"Data.Column1", type date}}),
    #"Added Custom1" = Table.AddColumn(#"Changed Type1", "Custom", each try Value.Is([Data.Column1], type date)
otherwise null),
    #"Filtered Rows1" = Table.SelectRows(#"Added Custom1", each ([Custom] = true)),
    #"Removed Columns1" = Table.RemoveColumns(#"Filtered Rows1",{"Custom"}),
    #"Added Custom" = Table.AddColumn(#"Removed Columns1", "Month", each //Text.Combine({Text.From(Date.Month([Data.Column1])),"/",Text.From(Date.Year([Data.Column1]))})//

let vuosi =

[
    luku = Text.From(Date.Year([Data.Column1])),
    kkName = Text.From(Date.MonthName([Data.Column1]))
]

in 
Text.Trim(Text.Clean(Text.Combine({vuosi[kkName]," ", vuosi[luku]})))),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom",{"Month", "Data.Column1", "Data.Column3", "Data.Column4"}),
    #"Merged Columns" = Table.CombineColumns(#"Reordered Columns",{"Data.Column5", "Data.Column4"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Description"),
    #"Renamed Columns" = Table.RenameColumns(#"Merged Columns",{{"Data.Column1", "Date"}, {"Data.Column3", "Accounting event"}}),
    #"Filtered Rows2" = Table.SelectRows(#"Renamed Columns", each ([Accounting event] <> null)),
    #"Added Custom3" = Table.AddColumn(#"Filtered Rows2", "Last day of the month", each Date.EndOfMonth([Date])),
    #"Replaced Value1" = Table.ReplaceValue(#"Added Custom3","−","-",Replacer.ReplaceText,{"Accounting event"}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","+","",Replacer.ReplaceText,{"Accounting event"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Replaced Value2",{{"Last day of the month", type date}, {"Accounting event", Currency.Type}}),
    #"Reordered Columns1" = Table.ReorderColumns( #"Changed Type2",{"Date", "Month", "Accounting event", "Description", "Last day of the month"}),
    BufferedValues = List.Buffer(#"Reordered Columns1"[Accounting event]),
    RT = Table.FromList(fxRunningTotal(BufferedValues), Splitter.SplitByNothing(), {"RT"}),
    Columns = List.Combine({Table.ToColumns(#"Reordered Columns1"), Table.ToColumns(RT)}),
    #"Converted to Table" = Table.FromColumns(Columns, List.Combine({Table.ColumnNames(#"Reordered Columns1"),{"Running Total"}})),
    #"Saldo added" = Table.AddColumn( #"Converted to Table", "Saldo", each List.Sum({XXXX, [Running Total]})), //XXXX IS THE INITIAL SALDO AT THE FIRST LINE OF RUNNING TOTAL//
    #"Saldo to currency" = Table.TransformColumnTypes(#"Saldo added",{{"Saldo", Currency.Type}})

in
   #"Saldo to currency"