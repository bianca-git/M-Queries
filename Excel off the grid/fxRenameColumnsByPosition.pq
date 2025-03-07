(Table as table, ColumnPositionList as list, NewColumnNamesList as list) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnPositionsList (list of numbers) - List of column numbers (e.g. {0,2}) 
- NewColumnNamesList (list of text) - List of new column names (e.g. {"NewColumn1", "NewColumn2", "NewColumn3"})

PURPOSE:
Renames columns based on their position (zero based)

NOTES:
The position and new column name must be in corresponding location in each list

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-07-13
VERSION: 1.1

*/
  let
    

    //Get list of Table column names based on position
    ColumnNamesList = List.Transform( ColumnPositionList,  each Table.ColumnNames(Table){ _ } ),

    //Rename Table columns if possible. If error do nothing
    RenameTables = Table.RenameColumns(Table, List.Zip({ColumnNamesList,NewColumnNamesList}), MissingField.Ignore)

  in
    RenameTables
