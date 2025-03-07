(Table as table, optional ColumnNamesList as list, optional ListTypeIsInclude as logical) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- [ColumnNamesList] (list of text) - List of column names to clean (e.g. {"Column1","Column2"})
- If excluded apply to all columns
- [ListTypeIsInclude] (logical) - Switch to determine if ColumnNamesList includes or excludes the list of names to clean.
 - true = apply to items in the list
 - false = apply to items not in the list

PURPOSE:
Changes columns names to new names based on rules. Inserts space and applies capitalization when a string:
- Changes from numbers to text
- Changes from text to numbers
- Changes from lower case to upper case
- Includes an underscore

NOTES:
Only use ColumnNamesList and ListTypeIsIncude where the action is to include or exclude specific column headers.
If used, both optional arguments are required.

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-04-05
VERSION: 1.1
*/

  let

    //Apply the Table.TransfromColumnNames transformation
    Result = Table.TransformColumnNames(
      Table, 
      (ColumnName as text) as text =>

        //Sub function to clean all column names
        let

          //Replace underscore
          replaceUnderScore = Text.Replace(ColumnName, "_", " "),

          //Apply Text.Clean to column name
          replaceSpecialCharacters = Text.Clean(ColumnName),

          //Split text on lower to upper case transition
          splitAtUpper = Splitter.SplitTextByCharacterTransition({"a" .. "z"}, {"A" .. "Z"})(
            replaceSpecialCharacters
          ), 
          combineSplitAtUpper = Combiner.CombineTextByDelimiter(" ", QuoteStyle.None)(splitAtUpper), 

          //Split text on text to number transition start
          splitAtNumberStart = Splitter.SplitTextByCharacterTransition(
            {"0" .. "9"}, 
            (c) => not List.Contains({"0" .. "9"}, c)
          )(combineSplitAtUpper), 
          combineSplitAtNumberStart = Combiner.CombineTextByDelimiter(" ", QuoteStyle.None)(
            splitAtNumberStart
          ), 

          //Split text on text to number transition end
          splitAtNumberEnd = Splitter.SplitTextByCharacterTransition(
            (c) => not List.Contains({"0" .. "9"}, c), 
            {"0" .. "9"}
          )(combineSplitAtNumberStart), 
          combineSplitAtNumberEnd = Combiner.CombineTextByDelimiter(" ", QuoteStyle.None)(
            splitAtNumberEnd
          ), 

          //Apply Proper Case
          returnValue = Text.Proper(combineSplitAtNumberEnd), 

          //Check if ColumnNamesList is provided.
          //If no list, return clean name, or if on list return clean name
          finalVal = 
            if ColumnNamesList is null then
              returnValue
            else if List.Count(List.FindText(ColumnNamesList, ColumnName))
              >= 1 and ListTypeIsInclude = true
            then
              returnValue
            else if List.Count(List.FindText(ColumnNamesList, ColumnName))
              = 0 and ListTypeIsInclude
              = false
            then
              returnValue
            else
              ColumnName

        in
          finalVal
    )

  in
    Result
