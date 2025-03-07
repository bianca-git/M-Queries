(Table as table, ColumnName as text, FilterList as list, optional ListIsInclude as logical) as table =>
/*
ARGUMENTS:
- Table (table) - Table or Step name to perform the transformation on
- ColumnName (text) - Name of column containing values
- FilterList (list) - List to filter by (e.g. {"Alpha","Bravo","Charlie"})
- ListIsInclude (logical) optional - Does the list exclude or include the items in the list.
    - true / null: Include the items in the list
    - false: Exclude the items in the list

PURPOSE:
Filters a table column based on a list

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-07-19
VERSION: 1.1

*/


let

    //Process parameters
    Temp1 = "Temp`¬|@#~1",

    //Get data type from first value in the column
    DataType = Value.Type(Table.Column(Table,ColumnName){0}),

    //Invert the filter if include <> true
    FinalList = if ListIsInclude <> false then List.Distinct(FilterList) else List.RemoveItems(List.Distinct(Table.Column(Table,ColumnName)),FilterList),

    //Convert list to table
    TableList = Table.FromColumns({FinalList}),

    //Apply data type to list
    TableListWithDataType = Table.TransformColumnTypes(TableList,{{"Column1", DataType}}),
    
    //Merge Tables
    MergeTables = Table.NestedJoin(TableListWithDataType, {"Column1"}, Table, {ColumnName}, Temp1, JoinKind.LeftOuter),
    
    //Remove tables with zero records
    RemoveZeroRecordTables = Table.SelectRows(MergeTables, each Table.RowCount(Record.Field(_, Temp1)) <> 0),

    //Remove unnecessary columns
    RemoveColumns = Table.SelectColumns(RemoveZeroRecordTables,{Temp1}),

    //Re-expand the column
    ExpandColumn = Table.ExpandTableColumn(
          RemoveColumns, 
          Temp1, 
          Table.ColumnNames(Table.Combine(Table.Column(RemoveColumns, Temp1)))
        )
in
    ExpandColumn
