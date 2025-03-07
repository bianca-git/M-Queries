(Table as table, GroupLength as number, optional KeepColumnNamesList as list, optional NewColumnNamesList as list) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on.
- GroupLength (number) - Number of rows contained in each group of columns.
- [KeepColumnNamesList] (list) - Name of columns which are unchanged
- [NewColumnNamesList] (list) - List of names to apply to unpivoted columns. Default names applied if null.

PURPOSE:
Unpivots multiple rows presented in a repeating row interval.

NOTES:
- NewColumnNamesList must contain one name for the unpivoted column headers, plus names for each item in the group length.

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2024-10-07
VERSION: 1.2
*/

let

    //Process Parameters
    Temp = "Temp`¬|@#~1",
    Temp2 = "Temp`¬|@#~2",

    //Force KeepColumnNamesList to a list
    NewKeepColumnNamesList = if KeepColumnNamesList = null then {} else KeepColumnNamesList,

    //Get list of column names to unpivot
    UnPivotColumnNamesList = List.RemoveItems(Table.ColumnNames(Table),NewKeepColumnNamesList),

    //Add Index column
    AddIndex = Table.AddIndexColumn(Table,Temp2,1,1/GroupLength),

    //Round Down to create repeating pattern
    RoundDown = Table.TransformColumns(AddIndex,{{Temp2, Number.RoundDown, Int64.Type}}),

    //Group the the columns to Pivot
    GroupRows = Table.Group(RoundDown, NewKeepColumnNamesList & {Temp2}, {{Temp, each _, type table}}),
    //Remove columns not pivoted & transpose
    TransformNestedTable = Table.TransformColumns(GroupRows, {{Temp, each 
        let
            #"Removed Other Columns" = Table.SelectColumns(_,UnPivotColumnNamesList),
            #"Demote Headers" = Table.DemoteHeaders(#"Removed Other Columns"),
            #"Transposed Table" = Table.Transpose(#"Demote Headers")
        in
            #"Transposed Table"
        }}),

    //Expand columns with new column names, or default column names
    ExpandData = if NewColumnNamesList = null
        then
        Table.ExpandTableColumn(TransformNestedTable, Temp, Table.ColumnNames(Table.Column(TransformNestedTable,Temp){0}))
        else
    Table.ExpandTableColumn(TransformNestedTable, Temp, Table.ColumnNames(Table.Column(TransformNestedTable,Temp){0}), NewColumnNamesList),
    
    //Remove Index / Grouping Column
    RemoveIndexGroupingColumn = Table.RemoveColumns(ExpandData,{Temp2})
    
in
    RemoveIndexGroupingColumn
