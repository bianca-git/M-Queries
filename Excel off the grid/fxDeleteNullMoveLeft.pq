(Table as table, ColumnName as text) =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnName (text) - Column name containing the null values to delete

PURPOSE:
Deletes null values in a column and shifts the cells left to fill the gap.

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-12-12
VERSION: 1.0
*/


let
    //Process parameters
    Temp = "Temp`¬|@#~1x",
    ColumnNumber = List.PositionOf(Table.ColumnNames(Table),ColumnName),

    //Add column to the Table
    AddCustomColumn = Table.AddColumn(Table, Temp, each 
        let
            //Convert record to a list
            RecordToList = Record.ToList(_),
            
            //Remove item from the list
            ListRemoveItem = if RecordToList{ColumnNumber} = null then 
                List.RemoveRange(RecordToList,ColumnNumber) & {null} else
                RecordToList,
            
            //Convert back to Record
            RecordToTable = Record.FromList(ListRemoveItem,Table.ColumnNames(Table))
        in
            RecordToTable),
    
    //Remove other columns
    RemovePreviousColumns = Table.SelectColumns(AddCustomColumn,{Temp}),

    //Expand columns
    ExpandedColumn = Table.ExpandRecordColumn(RemovePreviousColumns, Temp, Table.ColumnNames(Table))

in
    ExpandedColumn
