(Table as table, ColumnNamesList as list) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- [ColumnNamesList] (list of text) – List of column name on which to remove spaces

PURPOSE:
Removes excess spaces from text values

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-04-29
VERSION: 1.2

*/
  let

    fxTextRemoveSpaces = (TextValue as text) =>
      let

        //Clean the text
        cleanText = Text.Clean(TextValue), 

        //Trim the text
        trimText = Text.Trim(cleanText), 

        //Replace nbsp characters
        replaceNbsp = Text.Replace(trimText, "#(00A0)", " "), 

        //Split the text at each space character
        SplitText = Text.Split(replaceNbsp, " "), 

        //Remove the blank items from the list
        ListNonBlankValues = List.Select(SplitText, each _ <> ""), 

        //Join the list with a space character between each item
        TextJoinList = Text.Combine(ListNonBlankValues, " ")

      in
        TextJoinList, 

    //List of column names to transform
    TransformColumnList = if ColumnNamesList = null then Table.ColumnNames(Table) else ColumnNamesList,

    ColumnCount = List.Count(TransformColumnList),

    //Clean the text by calling the function
    CleanText = Table.TransformColumns(
      Table,
      List.Zip({
        TransformColumnList,
        List.Repeat({each fxTextRemoveSpaces(_)},ColumnCount),
        List.Repeat({type text},ColumnCount)}
      )
    )

  in
    CleanText
