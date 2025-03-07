(Table as table, PivotByColumnName as text, PivotValuesColumnName as text) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- PivotByColumnName (text) - Name of column to pivot on
- PivotValuesColumnName (text) - Name of values column to include in the pivot

PURPOSE:
Pivots all rows without causing an error

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
INSPIRED BY: https://www.youtube.com/watch?v=Tzv7jZ9WjAM
DATE: 2023-05-31
VERSION: 1.1
*/

let

    //Process Parameters
    Temp = "Temp`¬|@#~x",//Data
    Temp1 = "Temp`¬|@#~1", //Index

    //Get column names
    ColumnNames = Table.ColumnNames(Table),

    //Remove Pivot Field
    ReducedColumnNames = List.RemoveItems(ColumnNames,{PivotValuesColumnName}),

    //Group The Rows
    GroupRows = Table.Group(Table, ReducedColumnNames, {{Temp, each _, type table}}),
    
    //Add Nested Index
    AddNestedIndex = Table.TransformColumns(GroupRows, {{Temp, each Table.AddIndexColumn(_, Temp1, 0, 1, Int64.Type), type table}}),
    
    //Remove Columns
    RemoveColumns = Table.SelectColumns(AddNestedIndex,{Temp}),
    
    //Expand Column
    ExpandColumn = Table.ExpandTableColumn(RemoveColumns, Temp, ColumnNames & {Temp1}),
    //Pivot By
    PivotData = Table.Pivot(ExpandColumn, List.Distinct(Table.Column(ExpandColumn,PivotByColumnName)), PivotByColumnName, PivotValuesColumnName),

    //Remove Index
    RemoveIndex = Table.RemoveColumns(PivotData,{Temp1})
    
in
    RemoveIndex
