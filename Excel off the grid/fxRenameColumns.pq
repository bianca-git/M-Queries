(Table as table, ColumnNamesList as list, NewColumnNamesList as list) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnNamesList (list of text) - List of columns to rename (e.g. {"Column1","Column2","Column3"})
- NewColumnNamesList (list of text) - List of new column names (e.g. {"NewColumn1", "NewColumn2", "NewColumn3"})

PURPOSE:
Renames a column, if the column name exists

NOTES:
The old and new column names must be in corresponding location in each list

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-07-13
VERSION: 1.1

*/


let
    //Rename Table columns if possible. If error do nothing
    RenameTables = Table.RenameColumns(Table, List.Zip({ColumnNamesList,NewColumnNamesList}), MissingField.Ignore)
    
in
    RenameTables
