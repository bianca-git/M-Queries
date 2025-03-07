(Table as table, DateColumnName as text, FinancialYearEndMonth as number, NewColumnName as text, TimePeriod as text) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- DateColumnName (text) - Name of the column containing dates
- FinancialYearEndMonth (number) - Month number of the financial year end. Must be whole number from 0 to 12
- NewColumnName (text) - Name of the new column
- TimePeriod (text) - Set return value as Financial Year, Month or Quarter
  - "Month": Returns the Financial Month
  - "Quarter": Return the Financial Quarter
  - "Year": Return the Financial Year

PURPOSE:
Adds a Financial Month, Year, or Quarter column.

NOTES:
  - Assumes financial periods end on the last day of each month

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-06-13
VERSION: 1.1
*/

let

AddMonthColumn = if TimePeriod = "Month" then 
//Add a column to the Table
    Table.AddColumn(Table, NewColumnName, (_) =>
        let

        //Calcualte the calendar month
        Month = Date.Month(Record.Field(_, DateColumnName)),

        //Calcualte the Month offset by the Financial Year End Month
        TempFinMonth = Month - FinancialYearEndMonth,

        //Readjust the Financial Year Month if less that 0
        FinancialMonth = if TempFinMonth <= 0 then TempFinMonth + 12 else TempFinMonth
        in

        //Return the value as a whole number
        FinancialMonth, Int64.Type)

    else 
        Table,

AddQuarterColumn = if TimePeriod = "Quarter" then 
//Add a column to the Table
    Table.AddColumn(AddMonthColumn, NewColumnName, (_) =>
        let

        //Calcualte the calendar month
        Month = Date.Month(Record.Field(_, DateColumnName)),

        //Calcualte the Month offset by the Financial Year End Month
        TempFinMonth = Month - FinancialYearEndMonth,

        //Readjust the Financial Year Month if less that 0
        FinancialMonth = if TempFinMonth <= 0 then TempFinMonth + 12 else TempFinMonth,
        
        //Calculate Quarter
        FinancialQuarter = Number.RoundUp(FinancialMonth / 3)
        
        in

        //Return the value as a whole number
        FinancialQuarter, Int64.Type)

    else 
        AddMonthColumn,

AddYearColumn = if TimePeriod = "Year" then 

    Table.AddColumn(AddQuarterColumn, NewColumnName, (_) =>
        let

        //Calcualte the calendar month
        Month = Date.Month(Record.Field(_, DateColumnName)),

        //Calculate the calendear year
        Year = Date.Year(Record.Field(_, DateColumnName)),

        //Calcualte the Month offset by the Financial Year End Month
        TempFinMonth = Month - FinancialYearEndMonth,

        //Readjust the Financial Year Month if less that 0
        FinancialYear = if TempFinMonth <= 0 then Year else Year + 1
        in

        //Return the value as a whole number
        FinancialYear, Int64.Type)
    
    else AddQuarterColumn


in
    AddYearColumn
