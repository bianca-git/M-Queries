(Table as table, optional Threshold as number, optional SampleSize as number) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- [Threshold] (number) - Declare the % of valid items before applying the data type (e.g. 0.75 - 75% of the data needs to be of a specific data type otherwise treated as 'any’).
  - If omitted, default value is 1 (e.g., 100%)
- [SampleSize] (number) - Declare how many records to include in the sample.
  - If omitted, default value is 10.

PURPOSE:
Auto applies data types based on the data in a column.

NOTES:
- On large data sets this transformation can be slow, so recommended to always run on a sample.
- Where SampleSize exceeds the number of rows, the full data set is used.
- Using a Threshold < 0 or > 1 results in all columns being 'any' data type.
- DateTime data types are converted to Date where the time value is 00:00:00
- Detects Whole Numbers, Decimal Numbers, Dates, DateTimes, Logical and Text. Others treated as as 'any' or as 'text'.

AUTHOR: Mark Proctor / Excel Off The Grid
INSPIRED BY: https://datachant.com/2018/05/14/automatic-detection-of-column-types-in-powerquery/
DATE: 2023-05-15
VERSION: 1.1
*/


let
    //Create the function to assess each column
    fxColumnDataType = (Table as table, ColumnName as text) =>

        let

        //Drill down into the column
        DrillDown = Table.Column(Table, ColumnName),

        //Set default values
        TempSampleSize = if SampleSize is null then 1  else SampleSize,
        TempThreshold = if Threshold is null then 1 else Threshold,

        
        //Get the column of values
        DataList = List.FirstN(DrillDown,TempSampleSize),

        //Count types of data
        WholeNumberCount = List.Count(List.Select(DataList , each _ is number and _ = Number.Round(_, 0))),
        NumberCount = List.Count(List.Select(DataList , each _ is number)),
        DateCount = List.Count(List.Select(DataList , each _ is date)),
        WholeDateCount = List.Count(List.Select(DataList , each _ is datetime and Number.From(_) = Number.Round(Number.From(_), 0))),
        DateTimeCount = List.Count(List.Select(DataList , each _ is datetime)),
        TimeCount = List.Count(List.Select(DataList , each _ is time)),
        LogicalCount = List.Count(List.Select(DataList , each _ is logical)),
        TextCount = List.Count(List.Select(DataList , each _ is text)),
        MaxCount = List.Max({WholeNumberCount, NumberCount, DateCount, DateTimeCount, TimeCount, LogicalCount, TextCount}),
        
        //Return list of column name and data type to return
        SelectedType =
                if MaxCount / List.Count (DataList) < TempThreshold then
                    {ColumnName, type any}
                else if LogicalCount = MaxCount then
                    {ColumnName, type logical}
                else if WholeDateCount = MaxCount then
                    {ColumnName, type date}
                else if DateTimeCount = MaxCount then
                    {ColumnName, type datetime}
                else if DateCount = MaxCount then
                    {ColumnName, type date}
                else if TimeCount = MaxCount then
                    {ColumnName, type time}
                else if WholeNumberCount = MaxCount then
                    {ColumnName, Int64.Type}
                else if NumberCount = MaxCount then
                    {ColumnName, type number}
                else if TextCount = MaxCount then
                    {ColumnName, type text}
                else
                    {ColumnName, type any}

    in
        SelectedType,

    //Create list of all column type transformations
    AutoConvertTypes =
        List.Transform(
            Table.ColumnNames(Table),
            each fxColumnDataType(Table, _)
        ),

    //Apply data types to the Table
    ApplyTypes = Table.TransformColumnTypes(Table, AutoConvertTypes)
in
    ApplyTypes
