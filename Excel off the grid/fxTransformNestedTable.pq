(Table as table, ColumnName as text, eachFunctionList as list ) as table =>
/*
ARGUMENTS:
- Table (table) - Table or Step containing the nested Tables
- ColumnName (text) - Name of column containing nested Tables
- eachFunctionList (text) - List of functions to perform on the nested Table
    - Every function must be preceeded by the word each
    - Instead of refering to the Table, use underscore ( _ ) to refer to each nested Table

PURPOSE:
Perform Table transformations on nested Tables

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-07-15
VERSION: 1.1
*/

let
    TransformList = List.Accumulate(
        {0..List.Count(eachFunctionList)-1},
        Table,
        (state,current) => Table.TransformColumns(state, {{ColumnName, eachFunctionList{current}, type table}})
    )

in
    TransformList
