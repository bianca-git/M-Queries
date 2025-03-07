(NamedRange as text, optional DataTypeName as text) =>
/*
ARGUMENTS:
- NamedRange (text) - Excel named range containing the parameter.
- [DataTypeName] (text) - Data type of the parameter value.
  - Valid values include "Text", "Number", "Date/Time", "Date", "Time", "Duration", "True/False"
  - If excluded, the default type is Any

PURPOSE:
Returns a value from a named range in the current workbook to use as a parameter inside a query.

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-05-17
VERSION: 1.1
*/

  let
    
    //Get the Excel Table
    Table = Excel.CurrentWorkbook(){[Name = NamedRange]}[Content], 
    
    //Record of data types
    DataTypes = [ 
      Text = type text,
      String = type text,
      Number = type number,
      #"Date/Time" = type datetime,
      Date = type date,
      Time = type time,
      Duration = type duration,
      #"True/False" = type logical,
      Logical = type logical,
      Boolean = type logical
    ],


    //Get the data type
    GetDataType = Record.FieldOrDefault (DataTypes, Text.Proper(DataTypeName), type any),

    //Change the data type of the value
    ChangeType = Table.TransformColumnTypes(Table, {{"Column1", GetDataType}}), 
    
    //Return the first value form the Table
    ReturnValue = Table.Column(ChangeType, "Column1"){0}
  
  in
    ReturnValue
