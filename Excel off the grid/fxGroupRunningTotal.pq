(Table as table, ColumnName as text, NewColumnName as text, GroupByColumnNamesList as list, optional RetainOrder as logical) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnName (text) - Column name to perform the calculation on
- NewColumnName (text) - Name of the new running total column
- GroupByColumnNamesList (list of text) - List of the column names to group by
- [RetainOrder] (logical) - Should data be presented in the original order
  - true: values are returned to their original order
  - false / null: values ordered within each group

PURPOSE:
Adds a running total column where the total resets at the start of each group.

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
INSPIRED BY: https://gorilla.bi/power-query/running-total-by-category/
DATE: 2024-03-11
VERSION: 1.2
*/
  let

    //Process parameters
    Temp1 = "Temp`¬|@#~1", 
    Temp2 = "Temp`¬|@#~2", 

    //Function to calculate running total
    fxRunningTotal = (innerTable as table, ColumnName as text, NewColumnName as text) =>
      let

        //Convert column to number
        ApplyDataType = Table.TransformColumnTypes(innerTable, {{ColumnName, type number}}), 

        //Buffer the list of values
        BufferList = List.Buffer(Table.Column(ApplyDataType, ColumnName)), 

        //Calcualte the running total values
        RunningTotal =  List.Generate ( 
          () => [ RT = BufferList{0}, RowIndex = 0 ],
            each [RowIndex] < List.Count( BufferList ),  
            each [    
                  RT = List.Sum( { [RT], BufferList{[RowIndex] + 1} } ), 
                  RowIndex = [RowIndex] + 1 
                ],            
            each [RT] 
            ), 

        //Final table
        FinalTable = Table.FromColumns(
          Table.ToColumns(innerTable) & {RunningTotal}, 
          Table.ColumnNames(innerTable) & {NewColumnName}
        ), 

        //Convert column to number
        ApplyFinalDataType = Table.TransformColumnTypes(FinalTable, {{NewColumnName, type number}})
      in
        ApplyFinalDataType, 

    //Force data type of the column for the running total to number
    ChangeType = Table.TransformColumnTypes(Table, {{ColumnName, type number}}), 

    //Add an index column to ensure order can be retained
    AddIndex = Table.AddIndexColumn(ChangeType, Temp2), 

    //Group the rows and include all rows in each table
    GroupRows = Table.Group(AddIndex, GroupByColumnNamesList, {{Temp1, each _, type table}}), 

    //Add the running total to each table in the column, using function above
    AddRunningTotal = Table.TransformColumns(
      GroupRows, 
      {{Temp1, each fxRunningTotal(_, ColumnName, NewColumnName)}}
    ), 

    //Remove any other columns
    RemoveOtherColumns = Table.SelectColumns(AddRunningTotal, {Temp1}), 

    //Expand the Table to include all the columns
    ExpandedData = Table.ExpandTableColumn(
      RemoveOtherColumns, 
      Temp1, 
      Table.ColumnNames(AddIndex) & {NewColumnName}
    ), 

    //If RetainOrder = true, then put into order, else do nothing
    ReSortData = 
      if RetainOrder = true then
        Table.Sort(ExpandedData, {{Temp2, Order.Ascending}})
      else
        ExpandedData, 

    //Remove the added index column
    RemoveIndexColumn = Table.RemoveColumns(ReSortData, {Temp2}), 

    //Create a table with the correct data types for each column. Value.Type requires a Table with the exact same number of columns.
    CreateTableWithDataTypes = Table.AddColumn(ChangeType, NewColumnName, each null, type number), 

    //Restore the data types into the final table using Value.ReplaceType
    RestoreDataTypes = Value.ReplaceType(RemoveIndexColumn, Value.Type(CreateTableWithDataTypes))

  in
    RestoreDataTypes
