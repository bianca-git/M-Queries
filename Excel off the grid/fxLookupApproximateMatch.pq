(Table as table, ValueColumnName as text, LookupTable as table, LookupValueColumnName as text, LookupReturnColumnName as text, NewColumnName as text, optional LessThanOrMoreThan as text) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ValueColumnName (text) - Name of the column containing the lookup value.
- LookupTable (table) - Table to lookup the value from
- LookupValueColumnName (text) - Name of the column to lookup
- LookupReturnColumnName (text) - Name of the column to return value from
- NewColumnName (text) - The name of the column to be added
- [LessThanOrMoreThan] (text) - Should lookup be less than or equal to, or more than or equal to the lookup value.
  - "LessThan" / [any other value] / null : Return values less than or equal to the value
  - "MoreThan": Return values more than or equal to the value

PURPOSE:
Returns an approximate match lookup from another table.

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-03-30
VERSION: 1.1
*/
  let

    //Process Parameters
    IndexColName = "Temp`¬|@#~1", 
    MatchColName = "Temp`¬|@#~2", 
    Ordering = if LessThanOrMoreThan = "MoreThan" then Order.Descending else Order.Ascending, 

    //Add an index column to the Fact Table - this is to give a reference for getting records back 
    //  into their original order
    addIndex = Table.AddIndexColumn(Table, IndexColName, 0, 1), 
    
    //Keep necessary columns in Lookup table
    lookupTableKeepColumns = Table.SelectColumns(
      LookupTable, 
      {LookupValueColumnName, LookupReturnColumnName}
    ), 
    
    //Rename the columns to match in the Lookup and Fact tables
    factTableRenameColumn = Table.RenameColumns(addIndex, {ValueColumnName, MatchColName}), 
    lookupTableRenameColumn = Table.RenameColumns(
      lookupTableKeepColumns, 
      {{LookupValueColumnName, MatchColName}, {LookupReturnColumnName, NewColumnName}}
    ), 
    
    //Append the two tables
    tablesAppend = Table.Combine({factTableRenameColumn, lookupTableRenameColumn}), 
    
    //Sort the Values in ascending order
    sortValueCol = Table.Sort(tablesAppend, {{MatchColName, Ordering}, {IndexColName, Ordering}}),

    //Buffer the sort to ensure the sort order is retained
    tableBuffer = Table.Buffer(sortValueCol), 
    
    //Fill down the values column
    fillDown = Table.FillDown(tableBuffer, {NewColumnName}), 
    
    //Remove null values from the Index column
    removeNull = Table.SelectRows(fillDown, (_) => Record.Field(_, IndexColName) <> null), 
    
    //Sort the Index in ascending order
    sortIndexCol = Table.Sort(removeNull, {{IndexColName, Order.Ascending}}), 
    
    //Buffer the sort to ensure the sort order is retained
    tableBuffer2 = Table.Buffer(sortIndexCol), 
    
    //Remove the Index column
    removeIndex = Table.RemoveColumns(tableBuffer2, {IndexColName}), 
    
    //Rename Value column back to original name
    resetColumnNames = Table.RenameColumns(removeIndex, {MatchColName, ValueColumnName})
  
  in
    resetColumnNames
