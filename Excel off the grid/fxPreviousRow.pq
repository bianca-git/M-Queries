(Table as table, ColumnName as text, NewColumnName as text, RowOffset as number, optional ReturnOptions as text) as table =>
/*
- Table (table) - Table or step to perform the transformation on
- ColumnName (text) - Name of column containing values
- NewColumnName (text) - The name for the new column
- RowOffset (number) - A number representing the rows to offset by
- [ReturnOptions] (Text) - Determine type of value returned
  - "Value" / null: Returns the previous value
  - "Variance": Returns the numeric variance
  - "Percentage": Returns the percentage variance

PURPOSE:
Calculates the number or percentage variance from a previous row.

NOTES:
- RowOffset can include positive or negative numbers to get rows after or before
- All missing values are shown as null

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2024-04-11
VERSION: 1.2
*/

  let

    //Process parameter
    Temp1 = "Temp`¬|@#~1", 

    //Create null list for length of offset
    nullList = List.Repeat({null}, Number.Abs(RowOffset)), 

    //Create a column offset by x
    offsetColumn = 
      if RowOffset > 0 then
        nullList & List.RemoveLastN(Table.Column(Table, ColumnName), RowOffset)
      else
        List.RemoveFirstN(Table.Column(Table, ColumnName), Number.Abs(RowOffset)) & nullList, 

    //Convert Table to columns
    tableToColumns = Table.ToColumns(Table) & {offsetColumn}, 

    //Rejoin columns
    joinColumns = Table.FromColumns(tableToColumns, Table.ColumnNames(Table) & {Temp1}), 

    //Apply data type to new column
    applyDataType = Table.TransformColumnTypes(
      joinColumns, 
      {{Temp1, Value.Type(Table.Column(joinColumns,ColumnName){0})}}
    ), 

    //Calculate the variance
    calculateVariance = 
      if ReturnOptions = "Variance" then
        Table.AddColumn(
          applyDataType, 
          NewColumnName, 
          (_) => Record.Field(_, ColumnName) - Record.Field(_, Temp1), 
          type number
        )
      else
        applyDataType, 

    //Calculate the percentage
    calculatePercentage = 
      if ReturnOptions = "Percentage" then
        Table.AddColumn(
          calculateVariance, 
          NewColumnName, 
          (_) => (Record.Field(_, ColumnName) - Record.Field(_, Temp1)) / Record.Field(_, Temp1), 
          Percentage.Type
        )
      else
        calculateVariance, 

    //Remove the previous value column
    cleanUp = if ReturnOptions <> "Value" and ReturnOptions <> null then
    Table.RemoveColumns(calculatePercentage, {Temp1})
    else 
     Table.RenameColumns(calculatePercentage,{{Temp1, NewColumnName}})
    

  in
    cleanUp
