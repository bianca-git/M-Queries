(Table as table, ValueColumnNameList as list, LookupTable as table, LookupValueColumnNameList as list, LookupReturnColumnName as text, NewColumnName as text, optional ReturnOptions as text) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ValueColumnNameList (list of text) - Name of the columns containing the lookup value as a list (e.g. {"Col1","Col2","Col3"} for multicolumn, or {"Col1"} for a single column)
- LookupTable (table) - Table to lookup the value from.
- LookupValueColumnNameList (list of text) - Name of columns to lookup (e.g. {"Col1","Col2","Col3"} for multicolumn, or {"Col1"} for a single column).
- LookupReturnColumnName (text) - Name of the column to return value from
- NewColumnName (text) - Name of the column to be added
- [ReturnOptions] (text) - Which item(s) should the lookup return.
  - "First": Return the first item
  - "Last": Return the last item
  - [any other value] / null: Return all items

PURPOSE:
Returns an exact match lookup from another table.

NOTES:
- ValueColumnNames and LookupValueColumnNames list must contain the same number of items in the same order

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-04-11
VERSION: 1.0
*/


  let

  //Merge the tables
  mergeTables = Table.NestedJoin(Table, ValueColumnNameList, LookupTable, LookupValueColumnNameList, NewColumnName, JoinKind.LeftOuter),

  //Determin how to handle duplicates with the Return Options
  handleDuplicates = if ReturnOptions = "First" then
      Table.TransformColumns(mergeTables,{{NewColumnName, each Table.FirstN(_,1)}})
    else if ReturnOptions = "Last" then
      Table.TransformColumns(mergeTables,{{NewColumnName, each Table.LastN(_,1)}})
    else
      mergeTables,

  //Expand the column and return the value
  expandColumn = Table.ExpandTableColumn(handleDuplicates,NewColumnName, {LookupReturnColumnName},{NewColumnName})

  in
    expandColumn
