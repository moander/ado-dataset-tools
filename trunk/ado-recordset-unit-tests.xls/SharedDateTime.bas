Attribute VB_Name = "SharedDateTime"
Option Explicit

Function DayBeginning(myDate As Date) As Date
  Dim strDate As String
  Dim y, m, d As String
  
  ' Need to calculate the end of the date
  y = CStr(year(myDate))
  m = Lpad(month(myDate), "0", 2)
  d = Lpad(day(myDate), "0", 2)
  ' Universal Date format is : "yyyy-mm-dd hh:mm:ss" this ensure no confusion on dd/mm/yy vs mm/dd/yy
  strDate = y & "-" & m & "-" & d & " 00:00:00"
  DayBeginning = CDate(strDate)
End Function

Function DayEnding(myDate As Date) As Date
  Dim strDate As String
  Dim y, m, d As String
  
  ' Need to calculate the end of the date
  y = CStr(year(myDate))
  m = Lpad(month(myDate), "0", 2)
  d = Lpad(day(myDate), "0", 2)
  ' Universal Date format is : "yyyy-mm-dd hh:mm:ss" this ensure no confusion on dd/mm/yy vs mm/dd/yy
  strDate = y & "-" & m & "-" & d & " 23:59:59"
  DayEnding = CDate(strDate)
  
End Function


Function EndOfMonth(myDate As Date) As Date
  Dim strDate As String
  Dim yr As Integer
  Dim mth As Integer
  Dim y, m, d As String
  
  ' The end of the month CAN NOT be worked out by day zero of the next month
  yr = year(myDate)
  mth = month(myDate)
  
  If mth = 12 Then
    yr = yr + 1
    mth = 1
  Else
    mth = mth + 1
  End If
  
  y = CStr(yr)
  m = Lpad(mth, "0", 2)
  d = "01"
  
  ' Universal Date format is : "yyyy-mm-dd hh:mm:ss" this ensure no confusion on dd/mm/yy vs mm/dd/yy
  strDate = y & "-" & m & "-" & d & " 23:59:59"
  
  ' Since you have the 1st of the next month, subtract one day.
  EndOfMonth = CDate(strDate) - 1
  
End Function


Function DaysDuration(startDate As Date, finishDate As Date) As Long
  ' only works out how may days need to be considered. So if one minute of work is done then
  ' one day is returned.
  DaysDuration = Ceiling(finishDate - startDate)
End Function



Function DateToUniversal(ByVal myDate As Date) As String
  ' Universal Date format is : "yyyy-mm-dd hh:mm:ss" this ensure no confusion on dd/mm/yy vs mm/dd/yy
  DateToUniversal = Format(myDate, "yyyy-mm-dd Hh:Nn:Ss")


End Function

Function UniversalToDate(ByVal myString As String) As Date
  ' Universal Date format is : "yyyy-mm-dd hh:mm:ss" this ensure no confusion on dd/mm/yy vs mm/dd/yy
  ' VBA should be able to always correctly deal with this
  UniversalToDate = CDate(myString)

End Function

Function DoubletoNearestFraction(ByVal dbl As Double, ByVal nearestFraction As Double) As Double
  ' The purpose of this is to take a double and round it to the nearest fraction.
  ' EG: To work out the nearest quarter hour? DoubletoNearestFraction(dblHours,0.25)
  ' EG: To work out the nearest 5 minutes? DoubletoNearestFraction(dblHours,5/60)
  
  Dim fraction As Double
  Dim rounded As Double
  
  
  If nearestFraction = 0 Then
    DoubletoNearestFraction = dbl
    Exit Function
  End If
  
  'fraction = 1 / nearestFraction
  rounded = Round(dbl / nearestFraction, 0) * nearestFraction
  
  DoubletoNearestFraction = rounded


End Function
