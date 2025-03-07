(Table as table, ColumnNamesList as list, FindList as list, ReplaceList as list, optional PartialMatch as logical) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- ColumnNamesList (list of text) - List of column names to perform the find and replace action on
- FindList (list of text) - List of text to find (e.g. {"A","B","C"})
- ReplaceList (list of text) - List of text to replace (e.g. {"D","E","F"})
- PartialMatch (logical) - Can the find and replace be a partial mach
    - true: Find and replace is performed on a partial text string
    - false: Find and replace is performed on entire cell values only

PURPOSE:
Find and replace values based on a list

NOTES:
- The find and replace values must be in corresponding location in each list
- Transformation is canse sensitive

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-07-29
VERSION: 1.0
*/

let

    PartialMatchReplacer = if PartialMatch = true then Replacer.ReplaceText else Replacer.ReplaceValue,

    //Loop through multiple find and replace action
    MultipleFindReplace = List.Accumulate(
    {0..List.Count(FindList)-1},
    Table,
    (state, current) => 
            Table.ReplaceValue(
                state,
                FindList{current},
                ReplaceList{current},
                PartialMatchReplacer,
                ColumnNamesList)
    )
in
    MultipleFindReplace
