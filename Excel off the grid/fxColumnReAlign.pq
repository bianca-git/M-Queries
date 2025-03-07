(Table as table) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on

PURPOSE:
Re-align columns where data does not stack correctly in a single column.

NOTES:
- Re-alignment works where empty values are null values, and there are no genuine null values in the data set.

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-04-26
VERSION: 1.1
*/

  let

    //Process Parameters
    Spacer = "Temp`¬|@#~1", 
    UnlikelyChar = "Temp`¬|@#~2", 

    //Convert all columns to text data types
    makeAllText = Table.TransformColumnTypes(
      Table, 
      List.Transform(Table.ColumnNames(Table), each {_, type text})
    ), 

    //Replace all null values with Spacer
    replaceAllNulls = Table.ReplaceValue(
      makeAllText, 
      null, 
      Spacer, 
      Replacer.ReplaceValue, 
      Table.ColumnNames(makeAllText)
    ), 

    //Merge the columns with UnlikelyChar
    mergeColumns = Table.CombineColumns(
      replaceAllNulls, 
      Table.ColumnNames(replaceAllNulls), 
      Combiner.CombineTextByDelimiter(UnlikelyChar, QuoteStyle.None), 
      "Data"
    ), 

    //Replace instances of Spaces & UnlikelyChar with space
    replaceValues = Table.ReplaceValue(
      mergeColumns, 
      Spacer & UnlikelyChar, 
      " ", 
      Replacer.ReplaceText, 
      {"Data"}
    ), 

    //Split by UnlikelyChar
    delimiterSplit = Table.SplitColumn(
      replaceValues, 
      "Data", 
      Splitter.SplitTextByDelimiter(UnlikelyChar, QuoteStyle.Csv)
    ), 

    //Trim the spaces
    trimText = Table.TransformColumns(
      delimiterSplit, 
      List.Transform(Table.ColumnNames(delimiterSplit), each {_, Text.Trim, type text})
    ), 

    //Convert blanks (i.e. Spacers) to null
    replaceBlanks = Table.ReplaceValue(
      trimText, 
      Spacer, 
      null, 
      Replacer.ReplaceValue, 
      Table.ColumnNames(trimText)
    ), 

    //Get the Table profile with the null count
    TableProfile = Table.Profile(replaceBlanks)[[Column],[NullCount]],

    //Get the list of null values
    ColumnList = Table.SelectRows(TableProfile, each ([NullCount] = Table.RowCount(replaceBlanks)))[Column],

    //Remove the null values
    RemoveColumns = Table.RemoveColumns(replaceBlanks,ColumnList)

  in
    RemoveColumns
