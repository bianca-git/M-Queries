(optional Unit as text) as table =>
/*
ARGUMENTS:
- [Units] (text) - Text value of time units:
  - "Hours": Hour time segments
  - "Minutes": Minute time segments
  - "Seconds" / [any other value] / null: Second time segments

PURPOSE:
Creates a time table based in hour, minute, or second intervals

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-07-07
VERSION: 1.1

*/
  let

    //Calcualte the units
    UnitNumber = if Unit = "Hours" then 24 else if Unit = "Minutes" then 24 * 60 else 24 * 60 * 60, 
    
    //Create a list the length of the units
    UnitList = {0 .. UnitNumber - 1},

    //Divide a day by the units
    DivideDayByUnits = List.Transform(UnitList, each _ / UnitNumber), 

    //Convert list ot Table
    ConvertListToTable = Table.FromList(
      DivideDayByUnits, 
      Splitter.SplitByNothing(), 
      type table [Time = Number.Type]
    ), 

    //Convert number to time data type
    ConvertToTime = Table.TransformColumnTypes(ConvertListToTable, {{"Time", type time}})
  
  in
    ConvertToTime
