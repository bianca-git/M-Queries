(TableName as text, ParameterName as text, ParameterNameColumn as text, ParameterValueColumn as text, optional DataTypeName as text) =>
/*
ARGUMENTS:
- TableName (text) - Table containing the parameter
- ParameterName (text) - Name of the parameter in the table
- ParameterNameColumn (text) - Column name containing the parameter names
- ParameterValueColumn (text) - Column name containing the value to return
- [DataTypeName] (text) - Data type of the parameter value.
  - Valid values include "Text", "Number", "Date/Time", "Date", "Time", "Duration", "True/False"
  - If excluded, the default type is Any

PURPOSE:
Returns a value from a table in the current workbook to use as a parameter inside a query.

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-05-17
VERSION: 1.1
*/

  let

    //Get the Excel Table
    Table = Excel.CurrentWorkbook(){[Name = TableName]}[Content], 

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

    //Filter Table to Parameter Name
    FilterTable = Table.SelectRows(Table, each Record.Field(_, ParameterNameColumn) = ParameterName), 

    //Get the data type
    GetDataType = Record.FieldOrDefault (DataTypes, Text.Proper(DataTypeName), type any),

    //Change the Parameter value to the the correct data type
    ChangeType = Table.TransformColumnTypes(FilterTable, {{ParameterValueColumn, GetDataType}}), 

    //Turn the Table into a single value
    ReturnValue = Table.Column(ChangeType, ParameterValueColumn){0}

  in
    ReturnValue
