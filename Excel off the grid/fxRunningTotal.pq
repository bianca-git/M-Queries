(Table as table, ColumnName as text, NewColumnName as text) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnName (text) - Column name to perform the calculation on
- NewColumnName (text) - Name of the new running total column

PURPOSE:
Adds a running total column

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
BASED ON: https://gorilla.bi/power-query/running-total/ by Rick De Groot
DATE: 2024-03-11
VERSION: 1.2
*/

  let

    //Convert column to a number
    ApplyDataType = Table.TransformColumnTypes(Table, {{ColumnName, type number}}), 

    //Buffer the list of values
    BufferList = List.Buffer(Table.Column(ApplyDataType, ColumnName)), 

    //Calculate the running total values
    RunningTotal = List.Generate ( 
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
      Table.ToColumns(Table) & {RunningTotal}, 
      Table.ColumnNames(Table) & {NewColumnName}
    ), 

    //Convert column to number
    ApplyFinalDataType = Table.TransformColumnTypes(FinalTable, {{NewColumnName, type number}})

  in
    ApplyFinalDataType
