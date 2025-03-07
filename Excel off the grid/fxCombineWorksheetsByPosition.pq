(Table as table, ContentColumnName as text, Position as number, SheetOrTable as text, optional PromoteHeaders as logical, optional AutoExpand as logical) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step containing the files in the folder
- ContentColumnName (text) - Name of column containing the workbook binaries
- Position (number) - Number of the sheet or table to combine (zero-based)
- SheetOrTable (text) - Determine if objects to combine are sheets or tables
- [PromoteHeaders] - (logical) - Should headers be promoted
  - true: Promote headers
  - false / null: Do not promote headers
- [AutoExpand] (logical) - Should data should be expanded automatically (dynamically).
  - true: Expand data
  - false / null: Do not expand data

PURPOSE:
Combine worksheets from workbooks in a folder using sheet or table position

NOTES:
(None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-05-31
VERSION: 1.1
*/
  let

    //Process Parameters
    Temp1 = "Excel", 
    Temp2 = "Temp`¬|@#~2", 
    Temp3 = "Temp`¬|@#~3", 

    //Add column for the workbook
    AddExcelWorkbook = Table.AddColumn(
      Table, 
      Temp1, 
      each Excel.Workbook(Record.Field(_, ContentColumnName), PromoteHeaders)
    ), 

    //Apply nested transformation
    ApplyNestedTransform = Table.TransformColumns(
      AddExcelWorkbook, 
      {
        {
          Temp1, 
          each 
            let


              //Filter By Object
              FilterByObject = Table.SelectRows(_, each _[Kind] = SheetOrTable),

              //Add Index
              AddIndex = Table.AddIndexColumn(FilterByObject, Temp3, 0, 1 ),

              //Filter to Position
              FilterRows = Table.SelectRows(
                    AddIndex, 
                    each Record.Field(_, Temp3)
                        = Position
                    ),
                  
              //Rename columns
              SheetColumnRename = Table.RenameColumns(FilterRows, {{"Name", "Excel Object Name"}}), 
              
              //Remove excessive columns
              RemoveColumns = Table.SelectColumns(SheetColumnRename, {"Excel Object Name", "Data"}), 

              
              //Expand columns dynamic
              DynamicExpand = Table.ExpandTableColumn(
                RemoveColumns, 
                "Data", 
                Table.ColumnNames(Table.Combine(Table.Column(RemoveColumns, "Data")))
              )

            in
              DynamicExpand, 
          type table
        }
      }
    ), 
    
    //Remove any workbooks with zero records
    RemoveNullTables = Table.SelectRows(ApplyNestedTransform, each Table.RowCount(_[Excel]) > 0),

    //dynamically expand the data column dynamically
    ExpandData = 
      if AutoExpand = true then
        Table.ExpandTableColumn(
          RemoveNullTables, 
          Temp1, 
          Table.ColumnNames(Table.Combine(Table.Column(RemoveNullTables, Temp1)))
        )
      else
        ApplyNestedTransform
  
  in
    ExpandData
