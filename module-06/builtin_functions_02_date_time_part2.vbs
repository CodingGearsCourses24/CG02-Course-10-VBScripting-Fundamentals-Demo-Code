' Using built-in functions - Working with dates
' Now
' Date, DatePart, DateAdd, Weekday, Day, Month, MonthName, Year, Weekday, WeekdayName
' DateAdd, DateDiff, DataValue
' Time, Hour, Minute, Second, TimeValue, Time

mydate_set = #03-15-2015#
mytime = Time
myhour = Hour(MyTime)
myminute = Minute(MyTime)
mysecond = Second(MyTime)

'DisplayMessage mytime, "mytime"
'DisplayMessage myhour, "myhour"
'DisplayMessage myminute, "myminute"
'DisplayMessage mysecond, "mysecond"

' Weekday
weekday1 = Weekday(date)
weekday2 = WeekdayName(weekday1)
weekday3 = WeekdayName(Weekday(Now))

'DisplayMessage weekday1, "weekday1"
'DisplayMessage weekday2, "weekday2"
'DisplayMessage weekday3, "weekday3"

'IsDate
check_date1 = IsDate(mydate_set)
DisplayMessage check_date1, "check_date1"

Function DisplayMessage(message, id)
	MsgBox id & " : " & message,0,"Welcome"
End Function
