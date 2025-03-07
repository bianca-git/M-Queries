(StartDate as date, optional EndDate as date, optional DaysDuration as number) as table =>
/*
ARGUMENTS:
- StartDate (date) - The first date in the calendar
- [EndDate] (date) - The last date in the calendar
- [DaysDuration] (number) - The number of days from the start date

PURPOSE:
Creates a calendar table based on start date & end date, or start date & duration

NOTES:
- If EndDate is provided a calendar is created using StartDate & EndDate
- If EndDate is null a calendar is created using StartDate & DaysDuration

AUTHOR: Mark Proctor / Excel Off The Grid
DATE: 2023-04-25
VERSION: 1.1
*/


  let

    //Create the date list
    dateList = 
      if EndDate = null then
        List.Dates(StartDate, DaysDuration, #duration(1, 0, 0, 0))
      else
        List.Dates(StartDate, Duration.Days(EndDate - StartDate) + 1, #duration(1, 0, 0, 0)), 

    //Create the Table
    dateTable = Table.FromList(dateList, Splitter.SplitByNothing(), type table [Date = Date.Type])
  in
    dateTable
