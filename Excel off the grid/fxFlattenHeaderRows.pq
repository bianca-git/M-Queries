(Table as table, HeaderRowCount as number, Separator as text, optional FillDirection as text) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- HeaderRowCount (number) - Number of header rows to flatten
- Separator (text) - The character(s) to use a separator between the individual elements of the final header row
- [FillDirection] (text) - The direction to fill merged or null columns
  - "Right" = Fill to the right;
  - "Left" = Fill to the left;
  - Null = no fill

PURPOSE:
Flattens multiple header rows into a single header row. 

NOTES:
- The header rows must be in the first rows of the data; not part of an existing header row

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2024-04-29
VERSION: 1.2
*/

  let
    //Process Parameters 
    //Note: TempSeparator contains multiple spaces because these are trimmable
    TempSeparator = "                       ", 
    HeaderTemp = "Temp¬|@#~x1", 

    //Get the header rows
    GetHeaderRows = Table.FirstN(Table, HeaderRowCount), 

    //Transpose the header rows
    ColumnsToRows = Table.Transpose(GetHeaderRows), 

    //Fill columns up or down to populate null values
    FillUpDownBlanks = 
      if FillDirection = "Right" then
        Table.FillDown(ColumnsToRows, Table.ColumnNames(ColumnsToRows))
      else if FillDirection = "Left" then
        Table.FillUp(ColumnsToRows, Table.ColumnNames(ColumnsToRows))
      else
        ColumnsToRows, 

    //Change column names to text
    ChangeColumnsToText = Table.TransformColumnTypes(
      FillUpDownBlanks, 
      List.Transform(Table.ColumnNames(FillUpDownBlanks), each {_, type text})
    ), 

    //Merge all the header columns with the TempSeparator
    MergeColumns = Table.CombineColumns(
      ChangeColumnsToText, 
      Table.ColumnNames(ChangeColumnsToText), 
      Combiner.CombineTextByDelimiter(TempSeparator), 
      HeaderTemp
    ), 

    //Trim the columns to remove the Temp Separator from start and end of column headers
    TrimColumn = Table.TransformColumns(MergeColumns, {{HeaderTemp, Text.Trim, type text}}), 

    //In the middle of the column header, replace the TempSeparator with the HeaderTemp
    ReplaceSeparator = Table.ReplaceValue(
      TrimColumn, 
      TempSeparator, 
      Separator, 
      Replacer.ReplaceText, 
      {HeaderTemp}
    ), 

    //Transpose back to header row
    RowsToColumns = Table.Transpose(ReplaceSeparator),

    //Return origial column headers
    ReturnOriginalHeaders = Table.RenameColumns(
      RowsToColumns, 
      List.Zip({
        Table.ColumnNames(RowsToColumns),
        Table.ColumnNames(Table)})
      ),

    //Get data from the original Table
    RemoveHeaderRows = Table.RemoveFirstN(Table, HeaderRowCount), 

    //Append data
    FullData = Table.Combine({ReturnOriginalHeaders, RemoveHeaderRows}), 

    //Promote Headers
    FinalTable = Table.PromoteHeaders(FullData)

  in
    FinalTable
