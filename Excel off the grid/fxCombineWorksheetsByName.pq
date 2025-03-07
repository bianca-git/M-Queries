(Table as table, ContentColumnName as text, ObjectName as text, PartialMatch as logical, SheetOrTable as text, optional PromoteHeaders as logical, optional AutoExpand as logical) as table =>
  /*
ARGUMENTS:
- Table (table) - Table or step containing the files in the folder
- ContentColumnName (text) - Name of column containing the workbook binaries
- ObjectName (text) - Text string of the table or sheet name to combine
- PartialMatch (logical) - Determines if only partial match required (includes ignoring case)
  - true: Partial match permitted
  - false / null: Exact match only
- SheetOrTable (text) - Determine if objects to combine are sheets or tables
- [PromoteHeaders] (logical) - Should headers be promoted
  - true: Promote headers
  - false / null: Do not promote headers
  [AutoExpand] (logical) - Should data should be expanded automatically (dynamically).
  - true: Expand data
  - false / null: Do not expand data

PURPOSE:
Combine worksheets from workbooks in a folder using sheet or table names.

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-05-31
VERSION: 1.1
*/
  let

    //Process Parameters
    Temp1 = "Excel", 
    Temp2 = "Temp`¬|@#~1", 

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

              //Filter to Name
              FilterRows = 
                if PartialMatch = true then
                  Table.SelectRows(
                    _, 
                    each (
                      Text.Contains(Record.Field(_, "Name"), ObjectName, Comparer.OrdinalIgnoreCase)
                        and Record.Field(_, "Kind")
                        = SheetOrTable
                    )
                  )
                else
                  Table.SelectRows(
                    _, 
                    each (Record.Field(_, "Name") = ObjectName)
                      and Record.Field(_, "Kind")
                      = SheetOrTable
                  ), 
              
              //Rename columns
              SheetColumnRename = Table.RenameColumns(FilterRows, {{"Name", "Excel Object Name"}}), 
              
              //Remove excessive columns
              RemoveColumns = Table.SelectColumns(SheetColumnRename, {"Excel Object Name", "Data"}), 
              
              //Change data type for Excel Object Name
              ExcelObjectNameDataType = Table.TransformColumnTypes(
                RemoveColumns, 
                {{"Excel Object Name", type text}}
              ), 
              
              //Expand columns dynamic
              DynamicExpand = Table.ExpandTableColumn(
                ExcelObjectNameDataType, 
                "Data", 
                Table.ColumnNames(Table.Combine(Table.Column(ExcelObjectNameDataType, "Data")))
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
