(Table as table, ColumnName as text, Interval as number) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnName (text) - Name of column containing the stacked data
- Interval (number) - Length of repeating interval

PURPOSE:
Unstacks a column of data into separate columns based on a repeating numerical interval.

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-04-04
VERSION: 1.1
*/

  let

    //Process parameters
    Temp1 = "Temp`¬|@#~1", 
    Temp2 = "Temp`¬|@#~2", 
    ColumnNamePrefix = "Column", 

    //Add an index column
    addIndex = Table.AddIndexColumn(Table, Temp1, 0, 1, Int64.Type), 

    //Add an integer division column
    addIntegerDivision = Table.AddColumn(
      addIndex, 
      Temp2, 
      (_) => Number.IntegerDivide(Record.Field(_, Temp1), Interval), 
      Int64.Type
    ), 

    //Transform index to modulo
    transformIndexModulo = Table.TransformColumns(
      addIntegerDivision, 
      {{Temp1, each Number.Mod(_, Interval), type number}}
    ), 

    //Table column Names
    columnNames = Table.ColumnNames(transformIndexModulo), 

    //List of data types
    dataTypeList = List.Repeat({type text}, List.Count(columnNames)), 

    //Zipped list of columnNames + data types
    zippedList = List.Zip({columnNames, dataTypeList}), 

    //Add 1 to give the correct final Column numbers
    Add1ToIndexModulo = Table.TransformColumns(
      transformIndexModulo, 
      {{Temp1, each _ + 1, type number}}
    ), 

    //Change column types
    changeTypeToText = Table.TransformColumnTypes(Add1ToIndexModulo, zippedList), 

    //Add column prefix
    addColumnPrefix = Table.TransformColumns(
      changeTypeToText, 
      {{Temp1, each ColumnNamePrefix & _, type text}}
    ), 

    //Pivot without aggregating
    pivotData = Table.Pivot(
      addColumnPrefix, 
      List.Distinct(Table.Column(addColumnPrefix, Temp1)), 
      Temp1, 
      ColumnName
    ), 

    //Remove unrequired columns
    removeColumns = Table.RemoveColumns(pivotData, {Temp2})

  in
    removeColumns
