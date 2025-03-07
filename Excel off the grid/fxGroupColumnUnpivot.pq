(Table as table, GroupLength as number, optional KeepColumnNamesList as list, optional NewColumnNamesList as list) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on.
- Group Length (number) - Number of columns contained in each group of columns.
- [KeepColumnNamesList] (list) - Name of columns which are unchanged.
- [NewColumnNamesList] (list) - List of names to apply to unpivoted columns. Default column names applied if null.

PURPOSE:
Unpivots multiple columns presented in a repeating column interval.

NOTES:
- NewColumnNamesList must contain the same number of items as the repeating column pattern.

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2024-02-29
VERSION: 1.1
*/

let

    //Process Parameters
    Temp1 = "Temp`¬|@#~1", // Index
    Temp2 = "Temp`¬|@#~2", // Second Index
    Temp3 = "Temp`¬|@#~3", // Integer Division
    Temp4 = "Temp`¬|@#~4", // Modulo
    Temp5 = "Temp`¬|@#~5", // Attribute
    Temp6 = "Temp`¬|@#~6", // Value

    ColumnPrefix = "Column",

    //Add Index number to force a unique reference
    AddedIndex = Table.AddIndexColumn(Table, Temp1, 0, 1, Int64.Type),

    //KeepColumnNamesList or null
    NullKeepColumnNamesList = if KeepColumnNamesList = null then {} else KeepColumnNamesList,

    //Unpivot the columns
    UnpivotOtheColumns = Table.UnpivotOtherColumns(AddedIndex, NullKeepColumnNamesList & {Temp1}, Temp5, Temp6),

    //Remove columns, no longer required
    RemoveUnnecessaryColumns = Table.RemoveColumns(UnpivotOtheColumns,{Temp1, Temp5}),

    //Add Index column
    AddIndex2 = Table.AddIndexColumn(RemoveUnnecessaryColumns, Temp2, 0, 1, Int64.Type),

    //Convert Index to Grouped Index number
    GroupIndexNumber = Table.AddColumn(AddIndex2, Temp3, each Number.IntegerDivide(Record.Field(_, Temp2), GroupLength), Int64.Type),
    
    //Convert Index to Unpivot Grouped Index number
    UnpivotGroupIndexNumber = Table.AddColumn(GroupIndexNumber , Temp4, each Number.Mod(Record.Field(_, Temp2), GroupLength)+1, type number),

    //Remove Index column
    RemoveIndexColumn = Table.RemoveColumns(UnpivotGroupIndexNumber,{Temp2}),

    //Pivot on the Unpivot Grouped Index Number
    PivotColumn = Table.Pivot(Table.TransformColumnTypes(RemoveIndexColumn, {{Temp4, type text}}), List.Distinct(Table.Column(Table.TransformColumnTypes(RemoveIndexColumn, {{Temp4, type text}}),Temp4)), Temp4, Temp6),
    
    //Remove the Grouped Index Number
    RemoveGrouopIndexColumn = Table.RemoveColumns(PivotColumn,{Temp3}),
    
    //Get column names
    GetNewColumnNames = if NewColumnNamesList = null
        then List.Transform({"1"..Text.From(GroupLength)},each ColumnPrefix & _)
        else NewColumnNamesList,
    
    //Column Renames
    ColumnRenamesList = List.Zip({{"1"..Text.From(GroupLength)},GetNewColumnNames}),

    //Rename columns
    RenameColumns = Table.RenameColumns(RemoveGrouopIndexColumn,ColumnRenamesList)
in
    RenameColumns
