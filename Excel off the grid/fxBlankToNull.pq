(Table as table, optional ColumnNamesList as list) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- [ColumnNamesList] (list of text) - List of column names to convert blanks to null(e.g. {"Column1","Column2"}). If excluded defaults to all column names

PURPOSE:
Converts blank values in a table or column to null

NOTES:
 - (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-04-26
VERSION: 1.1
*/
  let

    //Convert Blank to Null
    convertBlankToNull = Table.ReplaceValue(
      Table, 
      "", 
      null, 
      Replacer.ReplaceValue,
      if ColumnNamesList = null then Table.ColumnNames(Table) else ColumnNamesList
    )
    
  in
    convertBlankToNull
