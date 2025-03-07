(Table as table, ColumnNamesList as list) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnNamesList (list of text) - List of column names containing the repeat values to change to null

PURPOSE:
Changes repeat values in a column to null

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2024-03-12
VERSION: 1.2
*/
  let

    unfillColumn = (Table, ColumnName) =>

      let
        
        //Process parameter
        Temp1 = "Temp`¬|@#~1", // Offset column
        Temp2 = "Temp`¬|@#~2", // Calculation column

        //New full list
        offsetColumn = {null} & List.RemoveLastN(Table.Column(Table, ColumnName), 1), 
       
        //Buffer the original column
        initialColumn = List.Buffer(Table.Column(Table, ColumnName)), 
        
        //Convert Table to columns
        tableToColumns = Table.ToColumns(Table) & {offsetColumn}, 
        
        //Rejoin columns
        joinColumns = Table.FromColumns(tableToColumns, Table.ColumnNames(Table) & {Temp1}), 

        //Add custom column
        customColumn = Table.AddColumn(joinColumns, 
          Temp2, 
          each if Record.Field(_, ColumnName) = Record.Field(_, Temp1) 
            then null 
            else Record.Field(_, ColumnName)),
        
        //Remove original columns
        removeOriginal = Table.RemoveColumns(customColumn , {ColumnName,Temp1}), 
        
        //Rename column
        renameColumn = Table.RenameColumns(removeOriginal, {{Temp2, ColumnName}}), 
       
        //Reorder columns
        reorderColumns = Table.ReorderColumns(renameColumn, Table.ColumnNames(Table))
      
      in
        reorderColumns,
    
    //Apply to multiple columns
    returnValue = List.Accumulate(
      ColumnNamesList, 
      Table, 
      (state, current) => unfillColumn(state, current)
    )
  in
    returnValue
