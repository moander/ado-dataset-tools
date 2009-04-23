Attribute VB_Name = "SharedDateTime"
Option Explicit

Function DayBeginning(myDate As Date) As Date
  Dim strDate As String
  Dim y, m, d As String
  
  ' Need to calculate the end of the date
  y = CStr(Year(myDate))
  m = Lpad(Month(myDate), "0", 2)
  d = Lpad(Day(myDate), "0", 2)
  ' Universal Date format is : "yyyy-mm-dd hh:mm:ss" this ensure no confusion on dd/mm/yy vs mm/dd/yy
  strDate = y & "-" & m & "-" & d & " 00:00:00"
  DayBeginning = CDate(strDate)


End Function

Function DayEnding(myDate As Date) As Date
  Dim strDate As String
  Dim y, m, d As String
  
  ' Need to calculate the end of the date
  y = CStr(Year(myDate))
  m = Lpad(Month(myDate), "0", 2)
  d = Lpad(Day(myDate), "0", 2)
  ' Universal Date format is : "yyyy-mm-dd hh:mm:ss" this ensure no confusion on dd/mm/yy vs mm/dd/yy
  strDate = y & "-" & m & "-" & d & " 23:59:59"
  DayEnding = CDate(strDate)
  
End Function


Function DaysDuration(startDate As Date, finishDate As Date) As Long
  ' only works out how may days need to be considered. So if one minute of work is done then
  ' one day is returned.
  DaysDuration = Ceiling(finishDate - startDate)

End Function

Function DateToUniversal_OLD(myDate As Date) As String
  Dim strDate As String

  Dim y, m, d As String
  Dim hh, mm, ss As String
  
  y = CStr(Year(myDate))
  m = Lpad(Month(myDate), "0", 2)
  d = Lpad(Day(myDate), "0", 2)
  
  hh = Lpad(Hour(myDate), "0", 2)
  mm = Lpad(Minute(myDate), "0", 2)
  ss = Lpad(Second(myDate), "0", 2)
  

  ' Universal Date format is : "yyyy-mm-dd hh:mm:ss" this ensure no confusion on dd/mm/yy vs mm/dd/yy
  strDate = y & "-" & m & "-" & d & " " & hh & ":" & mm & ":" & mm
  
 
  DateToUniversal_OLD = strDate

End Function


Function DateToUniversal(myDate As Date) As String
  ' Universal Date format is : "yyyy-mm-dd hh:mm:ss" this ensure no confusion on dd/mm/yy vs mm/dd/yy
  DateToUniversal = Format(myDate, "yyyy-mm-dd Hh:Nn:Ss")

End Function

Function UniversalToDate(myString As String) As Date
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
