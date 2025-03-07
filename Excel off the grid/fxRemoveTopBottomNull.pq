(Table as table, ColumnName as text, optional ReturnOptions as text) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnName (text) - Name of column to base transformation on
- [ReturnOptions] (text) - Which null values should be removed
    - "Top": Remove null values from the top
    - "Bottom": Remove null values from the bottom
    - [any other value] / null: Remove null values from Top and Bottom

PURPOSE:
Removes null values from top and/or bottom of a table

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-07-28
VERSION: 1.1
*/
  let

    //Process Parameters
    Temp = "Temp`¬|@#~x", 
    
    //Add Index column
    tableAddIndex = Table.AddIndexColumn(Table, Temp, 0, 1, Int64.Type), 
    
    //Filter table to remove null values
    filterTable = Table.SelectRows(
      tableAddIndex, 
      each Record.Field(_, ColumnName) <> null
    ),
    //Get index number from first and last row 
    firstRowNumber = List.Min(Table.Column(filterTable,Temp)),
    lastRowNumber = List.Max(Table.Column(filterTable,Temp)),
    
    //Remove top rows
    removeTopRows = if ReturnOptions <> "Bottom" then Table.Skip(Table, firstRowNumber) else Table,

    //Remove bottom rows
    removeBottomRows = if ReturnOptions <> "Top" then Table.RemoveLastN(removeTopRows, Table.RowCount(Table) - lastRowNumber - 1) else removeTopRows
  
  in
    removeBottomRows
