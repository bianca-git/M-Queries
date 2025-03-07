(Table as table, FilePathColumnName as text, DataColumnName as text, optional PromoteHeaders as logical, optional AutoExpand as logical) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step containing the nested tables
- ColumnName (text) - Name of column containing the nested tables
- eachFunctionList (List of nested functions) - List of functions to perform on the nested tables
  - Every function must be preceded by the word each
  - Refer to nested tables with an underscore ( _ )

PURPOSE:
Performs table transformations on nested tables

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-05-31
VERSION: 1.1
*/

  let
    //Process parameters
    Temp = "Excel Data Column", 
    
    //Add column with the data from the source
    addDataColumn = Table.AddColumn(
      Table, 
      Temp, 
      each Excel.Workbook(File.Contents(Record.Field(_, FilePathColumnName)), PromoteHeaders, true){
        [Item = Record.Field(_, DataColumnName)]
      }[Data]
    ), 

    //dynamically expand the data column dynamically
    expandData = 
      if AutoExpand = true then
        Table.ExpandTableColumn(
          addDataColumn, 
          Temp, 
          Table.ColumnNames(Table.Combine(Table.Column(addDataColumn, Temp)))
        )
      else
        addDataColumn

  in
    expandData
