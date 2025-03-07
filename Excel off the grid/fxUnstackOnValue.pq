(Table as table, ColumnName as text, SearchValue as text, optional IgnoreCase as logical) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnName (text) - Name of column containing the stacked data
- SearchValue (text) - Text value to split on
- [IgnoreCase] (logical) - Should case be ignored when matching the search value.
  - true: ignore case for the search,
  - false / null: apply case for search. 

PURPOSE:
Unstacks a column of data into separate columns based on the occurrence of a text string.

NOTES:
- Split occurs on Text values. To split on null or errors, replace null or errors with other values prior to unstacking

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-05-15
VERSION: 1.1

*/
  
  let
  
    //Process Parameters
    Temp1 = "Temp`¬|@#~1",  //INDEX
    Temp2 = "Temp`¬|@#~2",  //MATCH 
    Temp3 = "Temp`¬|@#~3",  //"Data"
  
    //Ignore case value
    IgnoreCaseValue = if IgnoreCase = true then Comparer.OrdinalIgnoreCase else Comparer.Ordinal, 
  
    //Add Index Column
    AddIndexColumn = Table.AddIndexColumn(Table, Temp1, 0, 1, Int64.Type), 
  
    //Get Index number if text matches
    GetIndexIfMatch = Table.AddColumn(
      AddIndexColumn, 
      Temp2, 
      each 
        if Text.Contains(Text.From(Record.Field(_, ColumnName)), SearchValue, IgnoreCaseValue) then
          Record.Field(_, Temp1)
        else
          null
    ), 
  
    //Remove any null errors
    ReplaceErrors = Table.ReplaceErrorValues(GetIndexIfMatch, {{Temp2, null}}), 
  
    //Fill Group Index Number down
    FillDownGroupIndex = Table.FillDown(ReplaceErrors, {Temp2}), 
  
    //Remove null values at the top
    RemoveNulls = Table.SelectRows(FillDownGroupIndex, each Record.Field(_, Temp2) <> null), 
  
    //Group by Group Index number
    GroupByIndex = Table.Group(RemoveNulls, {Temp2}, {{Temp3, each _, type table}}), 
  
    //Retain Search Column only
    RetainSearchColumn = Table.TransformColumns(
      GroupByIndex, 
      {{Temp3, each Table.SelectColumns(_, ColumnName), type table}}
    ), 
  
    //Transpose the grouped table
    TransposeGroupedTable = Table.TransformColumns(
      RetainSearchColumn, 
      {{Temp3, Table.Transpose, type table}}
    ), 
  
    //Retain column to expand in final table
    RetainDataColumn = Table.SelectColumns(TransposeGroupedTable, {Temp3}), 
  
    //Expand the data column
    ExpandDataColumn = Table.ExpandTableColumn(
      RetainDataColumn, 
      Temp3, 
      Table.ColumnNames(Table.Combine(Table.Column(RetainDataColumn, Temp3)))
    )
  
  in
    ExpandDataColumn
