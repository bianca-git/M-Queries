(Table as table, ColumnName as text, Units as text, UnitGroupSize as number, RoundType as text, NewColumnName as text) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnName (text) - Name of the column containing the time or datetime
- Units (text) - Text value of time units to group by:
  - "Hours": Hour time segments
  - "Minutes": Minute time segments
  - "Seconds" / [any other value] / null: Second time segments
- UnitGroupSize (number) - The length of each unit group  (e.g. When Units = "Minutes", 5 is 5 Minute segments)
- RoundType (text) - Text value of the how to round each value:
  - "Up": Round up
  - "Nearest": Round to nearest
  - "Down" / [any other value] / null: Round down
- NewColumnName (text) - The name of the column to be added

PURPOSE:
Groups time into Hour, Minute and Second time segments

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-07-07
VERSION: 1.1

*/
  let
    
    //Calculate total units into seconds
    TotalUnits = if Units = "Minutes" then 60 else if Units = "Hours" then 60 * 60 else 1, 
    
    //DataType of Source data
    GetDataType = Value.Type(Table.Column(Table, ColumnName){0}), 
    AddColumn = Table.AddColumn(
      Table, 
      NewColumnName, 
      each 
        let
          
          //Get the numger from the time
          NumberFromTime = Number.From([Time]), 
          
          //Calculate total units and round
          RoundedTime = 
            if RoundType = "Up" then
              Number.RoundUp(NumberFromTime * ((24 * 60 * 60) / (UnitGroupSize * TotalUnits)), 0)
            else if RoundType = "Nearest" then
              Number.Round(NumberFromTime * ((24 * 60 * 60) / (UnitGroupSize * TotalUnits)), 0)
            else
              Number.RoundDown(NumberFromTime * ((24 * 60 * 60) / (UnitGroupSize * TotalUnits)), 0), 
          
          //Divide units back to number 
          DivideToTime = RoundedTime / ((24 * 60 * 60) / (UnitGroupSize * TotalUnits)), 
          
          //Fix next day time issue
          NextDayFix = if DivideToTime = 1 then 0 else DivideToTime, 
          
          //Convert number back to time
          TimeType = 
            if GetDataType = type time then
              Time.From(NextDayFix)
            else
              DateTime.From(NextDayFix)
        
        in
          TimeType, 
      
      GetDataType
    )
  
  in
    AddColumn
