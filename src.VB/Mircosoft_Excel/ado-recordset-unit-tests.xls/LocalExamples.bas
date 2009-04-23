Attribute VB_Name = "LocalExamples"
Option Explicit


Sub Example_01()
  Dim sh As Worksheet
  Dim wb As Workbook
  Dim shName As String


  Dim rs As Recordset
  Dim i As Integer
  Dim txt As String
  Dim startDate As Date
  Dim endDate As Date
  
  endDate = UniversalToDate("2009-01-01 01:01:01") ' uses "yyyy-mm-dd hh:mm:ss"
  startDate = endDate - 90
  
  ' expect fields: Date,Open,High,Low,Close,Volume
  Set rs = GetGoogleFinanceData("NASDAQ:GOOG", startDate, endDate, , True, "The Google")
  
  Set wb = ActiveWorkbook
  shName = "Google EXAMPLE 01"
  Set sh = RecordsetToWorkSheet(rs, wb, shName)

End Sub

Sub Example_02()
  Dim sh As Worksheet
  Dim wb As Workbook
  Dim shName As String


  Dim rs, rs1, rs2, rs3 As Recordset
  Dim i As Integer
  Dim txt As String
  Dim startDate As Date
  Dim endDate As Date
  
  endDate = Now()
  startDate = endDate - 30
  
  ' expect fields: Date,Open,High,Low,Close,Volume
  Set rs1 = GetGoogleFinanceData("NASDAQ:GOOG", startDate, endDate, , True, "The Google")
  Set rs2 = GetGoogleFinanceData("NASDAQ:ORCL", startDate, endDate, True, True, "The Oracle") ' This returns only weekly data
  Set rs3 = GetGoogleFinanceData("NYSE:IBM", startDate, endDate, , , "The IBM")
  
  Set rs = MergeRecordset(rs1, rs2, rs3)
  
  Set wb = ActiveWorkbook
  shName = "Google EXAMPLE 02"
  Set sh = RecordsetToWorkSheet(rs, wb, shName)

End Sub


Sub Example_03()
  Dim sh As Worksheet
  Dim wb As Workbook
  Dim shName As String

  Dim rs As ADODB.Recordset
  Dim rs1, rs2, rs3 As ADODB.Recordset
  Dim i As Integer
  Dim txt As String
  Dim startDate As Date
  Dim endDate As Date
  
  endDate = Now()
  startDate = endDate - (365 * 1)
  
  ' expect fields: Date,Open,High,Low,Close,Volume
  Set rs1 = GetGoogleFinanceData("NASDAQ:GOOG", startDate, endDate, True, True, "The Google")
  Set rs2 = GetGoogleFinanceData("NASDAQ:ORCL", startDate, endDate, True, True, "The Oracle")
  Set rs3 = GetGoogleFinanceData("NYSE:IBM", startDate, endDate, True, True, "The IBM")
  
  Set rs = MergeRecordset(rs1, rs2, rs3)
  Set rs = AddFields(rs, "Month", , , adDate) ' specifying the new field to adDate helps excel set the format correctly
  rs.MoveFirst
  Do Until rs.EOF
    rs.Fields("Month").Value = EndOfMonth(rs.Fields("Date").Value)
    rs.MoveNext
  Loop
  
  Set rs = PivotRecordSet(rs, "CompanyName", "Month", "Close", "max", SORT_BY_GC)
  
  Set wb = ActiveWorkbook
  shName = "Google EXAMPLE 03"
  Set sh = RecordsetToWorkSheet(rs, wb, shName)

End Sub




Sub Example_04()
  Dim sh As Worksheet
  Dim wb As Workbook
  Dim shName As String

  Dim rs As ADODB.Recordset
  Dim rs1, rs2, rs3 As ADODB.Recordset
  Dim i As Integer
  Dim txt As String
  Dim startDate As Date
  Dim endDate As Date
  
  endDate = Now()
  startDate = endDate - (365 * 10)
  
  ' expect fields: Date,Open,High,Low,Close,Volume
  Set rs1 = GetGoogleFinanceData("NASDAQ:GOOG", startDate, endDate, True, True, "Google NASDAQ:GOOG")
  Set rs2 = GetGoogleFinanceData("NASDAQ:ORCL", startDate, endDate, True, True, "Oracle NASDAQ:ORCL")
  Set rs3 = GetGoogleFinanceData("NYSE:IBM", startDate, endDate, True, True, "IBM NYSE:IBM")
  
  Set rs = MergeRecordset(rs1, rs2, rs3)
  Set rs = AddFields(rs, "Month", , , adDate) ' specifying the new field to adDate helps excel set the format correctly
  rs.MoveFirst
  Do Until rs.EOF
    rs.Fields("Month").Value = EndOfMonth(rs.Fields("Date").Value)
    rs.MoveNext
  Loop
  
  Set rs = PivotRecordSet(rs, "CompanyName", "Month", "Close", "max", SORT_BY_GC)
  
  Set wb = ActiveWorkbook
  shName = "Google EXAMPLE 04"
  Set sh = RecordsetToWorkSheet(rs, wb, shName)

End Sub






