(TableList as list) as table =>
/*
ARGUMENTS:
- TableList (list of tables) - List of tables to join (e.g. {Table1,Table2,Table3})

PURPOSE:
Performs a cartesian join on all tables included a list. (i.e. creates a table of all combinations).

NOTES:
- TableList may also include names of previous steps

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-05-08
VERSION: 1.1
*/

  let

    //Process Parameters
    Temp1 = "Temp`¬|@#~1", 
    Temp2 = "Temp`¬|@#~2", 
    Temp3 = "Temp`¬|@#~3", 

    //Loop through each table in the TableList and perform the same merge transformation for each
    //Note: Counter starts at 1, as the first table is the opening state
    Result = List.Accumulate(
      {1 .. List.Count(TableList) - 1}, 
      TableList{0}, 
      (state, current) =>
        if current = 0 then
          state
        else
          let

            //Add a unique column name into Tbl1 with the value 1 in each row
            Tbl1AddTemp = Table.AddColumn(state, Temp1, each 1, Int64.Type), 

            //Get the column names from Tbl1
            Tbl1ColNames = Table.ColumnNames(Tbl1AddTemp), 

            //Add a unique column name into Tbl1 with the value 1 in each row
            Tbl2AddTemp = Table.AddColumn(TableList{current}, Temp2, each 1, Int64.Type), 

            //Get column names from second Table
            Tbl2ColNames = Table.ColumnNames(Tbl2AddTemp), 

            //Create unique column names
            Tbl2NewColNames = 
              let
                OriginalList     = Tbl1ColNames, 
                AdditionaList    = Tbl2ColNames, 
                temp1            = "Temp`¬|@#~1", 
                fullList         = List.Combine({OriginalList, AdditionaList}), 
                convertToTable   = Table.FromList(fullList, null, {temp1}), 
                transposeTable   = Table.Transpose(convertToTable), 
                promoteHeaders   = Table.PromoteHeaders(transposeTable, [PromoteAllScalars = true]), 
                fullColumnNames  = Table.ColumnNames(promoteHeaders), 
                clearColumnNmaes = List.RemoveFirstN(fullColumnNames, List.Count(OriginalList))
              in
                clearColumnNmaes, 

            //Merge Tbl2 into Tbl1 with using the 1 created in the steps above, with Full Outer join
            Merge = Table.NestedJoin(
              Tbl1AddTemp, 
              {Temp1}, 
              Tbl2AddTemp, 
              {Temp2}, 
              Temp3, 
              JoinKind.FullOuter
            ), 

            //Expand the merged columns
            Expand = Table.ExpandTableColumn(Merge, Temp3, Tbl2ColNames, Tbl2NewColNames), 

            //Remove the unique column names created at the start
            RemoveTemp = Table.RemoveColumns(Expand, {Temp1, Temp2})

          in
            RemoveTemp
    )

  in
    Result
