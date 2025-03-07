(Table as table, SearchColumnName as text, SearchValue as any, optional PartialMatch as logical, optional InstanceNumber as number) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- SearchColumnName (text) - Name of column to search
- SearchValue (any) - Value to search for
- [PartialMatch] (logical) - Determines if only partial match required (includes ignoring case)
    - true: Partial match permitted
    - false / null: Exact match only
- [InstanceNumber] (number): Where multiple matches exist, which instance to use as the basis for removing rows

PURPOSE:
Removes an unknown number of rows from the top until search term found

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-04-25
VERSION: 1.1
*/

  let
    
    //Process Parameters
    Temp = "Temp`¬|@#~x", 
    
    //Get instance number if null
    Instance = if InstanceNumber = null then 1 else InstanceNumber, 
    
    //Add Index column
    tableAddIndex = Table.AddIndexColumn(Table, Temp, 0, 1, Int64.Type), 
    
    //Filter table to find the SeachValue in the SearchColumnName
    filterTable = Table.SelectRows(
      tableAddIndex, 
      (_) =>
        if PartialMatch = true then
          Text.Contains(
            Text.From(Record.Field(_, SearchColumnName)), 
            Text.From(SearchValue), 
            Comparer.OrdinalIgnoreCase
          )
        else
          (Text.From(Record.Field(_, SearchColumnName)) = Text.From(SearchValue))
    ), 
    
    //Get index number from nth row
    matchRowNumber = Table.Column(filterTable, Temp){Instance - 1}, 
    
    //Remove rows above the match
    removeRows = 
      if Table.RowCount(filterTable) = 0 then
        Table.Skip(Table, Table.RowCount(Table))
      else
        Table.Skip(Table, matchRowNumber)
  
  in
    removeRows
