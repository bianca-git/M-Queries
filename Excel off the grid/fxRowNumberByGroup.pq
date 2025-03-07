(Table as table, GroupByColumnNamesList as list, NewColumnName as text, optional RetainOrder as logical) as table =>
/*
ARGUMENTS:
 - Table (table) - Table or step to perform the transformation on
 - GroupByColumnNamesList (list of text) - List of the column names to group by
 - NewColumnName (text) - Name of the column to add
 - [RetainOrder] (logical) - Should data be presented in the original order
    - true: values are returned to their original order
    - false / null: values ordered within each group

PURPOSE:
Adds a row number for each row in a group

NOTES: 
 - (none)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-10-23
VERSION: 1.1
*/

let

    Temp1 = "Temp`¬|@#~1",
    Temp2 = "Temp`¬|@#~2",
    Temp3 = "Temp`¬|@#~3",

    //Add the index number
    AddMasterIndex = Table.AddIndexColumn(Table,Temp2,1,1),

    //Group by the the GroupBy list
    GroupGroupBy = Table.Group(AddMasterIndex , GroupByColumnNamesList, {{Temp1, each _, type table}}),
    
    //Add the row index number
    AddIndex = Table.TransformColumns(GroupGroupBy, {{Temp1, each Table.AddIndexColumn(_,NewColumnName,1,1), type text}}),
    
    //Remove unrequired columns
    RemoveOtherColumns = Table.SelectColumns(AddIndex,{Temp1}),

    //Expand the GroupBy list
    DynamicExpand = Table.ExpandTableColumn(
     RemoveOtherColumns, 
     Temp1, 
     Table.ColumnNames(Table.Combine(Table.Column(RemoveOtherColumns, Temp1)))
    ),

    //Retain original order
    RetainOriginalOrder = if RetainOrder = true
        then Table.Sort(DynamicExpand,{{Temp2, Order.Ascending}})
        else DynamicExpand,

    //Remove original index column added
    RemoveMasterIndex = Table.RemoveColumns(RetainOriginalOrder,{Temp2}),

    //Create a table with the correct data types for each column. Value.Type requires a Table with the exact same number of columns.
    CreateTableWithDataTypes = Table.AddColumn(Table, NewColumnName, each null, Int64.Type), 

    //Restore the data types into the final table using Value.ReplaceType
    RestoreDataTypes = Value.ReplaceType(RemoveMasterIndex, Value.Type(CreateTableWithDataTypes))

in
    RestoreDataTypes
