' Using built-in functions - Working with dates
' Now
' Date, DatePart, DateAdd, Weekday, Day, Month, MonthName, Year, Weekday, WeekdayName
' DateAdd, DateDiff, DataValue
' Time, Hour, Minute, Second, TimeValue, Time


mydate_date = Date
mydata_now = Now
mydate_set = #03-15-2015#

month1 = DatePart("m", mydate_set)
day1 = DatePart("d", mydate_set)
year1 = DatePart("yyyy", mydate_set)
month1_name = MonthName(DatePart("m", mydate_set))

mynewdate1 = DateAdd("d", 2, mydate_set)
mynewdate2 = DateAdd("m", 3, mydate_set)

DisplayMessage mynewdate2, "mynewdate2 "

'DisplayMessage mydate_date, "mydate_date "
'DisplayMessage mydata_now, "mydata_now "
'DisplayMessage mydate_set, "mydate_set "
'DisplayMessage month1, "month1 "
'DisplayMessage day1, "day1 "
'DisplayMessage year1, "year1 "
'DisplayMessage month1_name, "month1_name "


Function DisplayMessage(message, id)
	MsgBox id & " : " & message,0,"Welcome"
End Function
