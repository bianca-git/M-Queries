(Table as table) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on

PURPOSE:
Removes any columns with 100% null values.

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2024-03-01
VERSION: 1.2
*/
  
  let
  
    //Get the Table profile with the null count
    TableProfile = Table.Profile(Table)[[Column],[NullCount]],

    //Get the list of null values
    ColumnList = Table.SelectRows(TableProfile, each ([NullCount] = Table.RowCount(Table)))[Column],

    //Remove the null values
    RemoveColumns = Table.RemoveColumns(Table,ColumnList)
    
  in
    RemoveColumns
