(Table as table, PeriodEndDateColumnName as text, NewDateColumnName as text) as table =>
/*
ARGUMENTS
- Table (table) - Table containing data about the period ends
- PeriodEndDateColumnName (text) - Name of column containing the period end date
- NewDateColumnName (text) - Name to give to the new Date column

PURPOSE:
Creates a calendar table from a table on non-calendar periods

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-07-29
VERSION: 1.1
*/

let

    //Process Parameters
    Temp1 = "Temp`¬|@#~x", 
    StartDate = List.Min(Table.Column(Table,PeriodEndDateColumnName)),
    EndDate = List.Max(Table.Column(Table,PeriodEndDateColumnName)),

    //Create the date list
    dateList = List.Dates(StartDate, Duration.Days(EndDate - StartDate) + 1, #duration(1, 0, 0, 0)), 

    //Create the Table
    dateTable = Table.FromList(dateList, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    
    //Rename column
    renameDateColumn = Table.RenameColumns(dateTable,{{"Column1",NewDateColumnName}}),

    //Change data type
    changeDataType = Table.TransformColumnTypes(renameDateColumn,{{NewDateColumnName, type date}}),

    //Merge the calendar table with the source table
    mergeDates = Table.NestedJoin(changeDataType, {NewDateColumnName}, Table, {PeriodEndDateColumnName}, Temp1, JoinKind.LeftOuter),

    //Expand the source table
    expandSource = Table.ExpandTableColumn(mergeDates, Temp1 , Table.ColumnNames(Table)),

    //Fill Up
    FillDataUp = Table.FillUp(expandSource,Table.ColumnNames(Table))

in
    FillDataUp
