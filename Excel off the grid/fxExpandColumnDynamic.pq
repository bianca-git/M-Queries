(Table as table, ColumnName as text, optional PromoteHeaders as logical) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnName (text) - Name of column containing the tables to expand
- [PromoteHeaders] (logical) - Should headers be promoted prior to expanding.
  - true: Promote headers
  - false / null: Do not promote headers

PURPOSE:
Expands a column dynamically including adding any new data

NOTES:
- Use fxAutoDetectDataType to automatically apply data type to columns

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-06-16
VERSION: 1.1
*/
  
  let
    
    //PromoteHeaders if requested
    Headers = if PromoteHeaders then
     Table.TransformColumns(Table, {{ColumnName, each Table.PromoteHeaders(_), type table}})
     else Table,


    //dynamically expand the data column dynamically
    expandData = Table.ExpandTableColumn(
      Headers, 
      ColumnName, 
      Table.ColumnNames(Table.Combine(Table.Column(Headers, ColumnName)))
    )
  
  in
    expandData
