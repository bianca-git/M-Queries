(Table as table, optional FillRight as logical) as table =>
/*
ARGUMENTS:
- Table (table) - Table or step to perform the transformation on
- [fillRight] (logical) - Determine direction to fill from
  - true: fill values to the right
  - false / null: fill values to the left

PURPOSE:
Fills null values with the values from the left or right

NOTES:
- (None)

AUTHOR: Mark Proctor / Excel Off The Grid
BASED ON: Melissa de Korte | https://www.linkedin.com/feed/update/urn:li:activity:7120778640262324226/
DATE: 2023-12-12
VERSION: 1.1
*/

let

    //Process Parameters
    Temp1 = "Temp`¬|@#~1",

    //Add column with record for each row where FillUp/FillDown applied to replace nulls
    Fill = if FillRight = false
        then Table.AddColumn(Table, Temp1, each
            Record.FromTable(Table.FillUp(Record.ToTable(_), {"Value"})))
        else Table.AddColumn(Table, Temp1, each
            Record.FromTable(Table.FillDown(Record.ToTable(_), {"Value"}))),

    //Restore data types
    RestoreDataTypes = Value.ReplaceType(Table.FromRecords(Table.Column(Fill,Temp1)), Value.Type(Table))

in 
    RestoreDataTypes
