Attribute VB_Name = "LocalUnitTests"
Option Explicit

' This module tests the following functions

' CloneRecordset()
' SortRecordSet()
' GroupRecordSet()
' PivotRecordSet()
' SubTotalRecordSet()
' PivotAndSubTotalRecordSet()
' AppendRecordset()
' MergeRecordset()

' If there is an Error in the sheet there is problems
' Merging adVariant to adVarChar
Sub RunAllTests()
  DeleteAllTestSheets
  
  Test_CloneRecordset_01
  Test_CloneRecordset_02
  
  Test_MergeRecordset_01
  Test_MergeRecordset_02
  Test_MergeRecordset_03
  
  Test_GroupRecordset_01
  Test_GroupRecordset_02
  Test_GroupRecordset_03
  Test_GroupRecordset_04
  
  Test_PivotRecordset_01
  Test_PivotRecordset_02
End Sub

Sub DeleteAllTestSheets()
   Dim sh As Worksheet
   Dim wb As Workbook
   
   Set wb = ActiveWorkbook
   
   Application.DisplayAlerts = False
   For Each sh In wb.Worksheets
     If sh.name Like "* TEST *" Then
       sh.Delete
     End If
   Next
   Application.DisplayAlerts = True
End Sub



Sub Test_CloneRecordset_01()
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long

   Dim rs, rs1, rs2, rs3, rs4 As Recordset
   
   Set wb = ActiveWorkbook
   shName = "CloneRecordSet TEST 01"

   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"), adVarChar, VARCHAR_SIZE)
   Set rs2 = CloneRecordset(rs1)

   cBefore = rs1.RecordCount + rs2.RecordCount
   
   
   rs2.MoveLast
   rs2.Delete

   
   cAfter = rs1.RecordCount + rs2.RecordCount
   Set rs = MergeRecordset(rs1, rs2)
   

   'MsgBox "Records before/after MergeRecordst() " & cBefore & "/" & cAfter & " " & rs.RecordCount
   Set sh = RecordSetToWorkSheet(rs, wb, shName)
End Sub


Sub Test_CloneRecordset_02()
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long

   Dim rs As ADODB.Recordset
   Dim rs1 As ADODB.Recordset
   Dim rs2 As ADODB.Recordset
   
   Set wb = ActiveWorkbook
   shName = "CloneRecordSet TEST 02"

   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"), adVariant)
   Set rs2 = CloneRecordset(rs1)

   cBefore = rs1.RecordCount + rs2.RecordCount
   
   rs2.MoveLast
   rs2.Delete
   
   cAfter = rs1.RecordCount + rs2.RecordCount

   

  'MsgBox "Records before/after MergeRecordst() " & cBefore & "/" & cAfter
   Set sh = RecordSetToWorkSheet(rs2, wb, shName)
End Sub


Sub Test_MergeRecordset_01()
   
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long

   Dim rs, rs1, rs2, rs3, rs4 As Recordset
   
   Set wb = ActiveWorkbook
   shName = "MergeRecordSet TEST 01"
   
   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"), adVarChar, VARCHAR_SIZE)
   Set rs2 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q2"), adVarChar, VARCHAR_SIZE)
   Set rs3 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q3"), adVarChar, VARCHAR_SIZE)
   Set rs4 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q4"), adVarChar, VARCHAR_SIZE)
   cBefore = rs1.RecordCount + rs2.RecordCount + rs3.RecordCount + rs4.RecordCount
   
   Set rs = MergeRecordset(rs1, rs2, rs3, rs4)
   cAfter = rs.RecordCount
   
   'MsgBox "Records before/after MergeRecordst() " & cBefore & "/" & cAfter
   
   Set sh = RecordSetToWorkSheet(rs, wb, shName)
   
End Sub

Sub Test_MergeRecordset_02()
   
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long

   Dim rs, rs1, rs2, rs3, rs4 As Recordset
   
   Set wb = ActiveWorkbook
   shName = "MergeRecordSet TEST 02"

   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"))
   Set rs2 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q2"))
   cBefore = rs1.RecordCount + rs2.RecordCount
   
   Set rs = MergeRecordset(rs1, rs2)
   cAfter = rs.RecordCount
   
   'MsgBox "Records before/after MergeRecordst() " & cBefore & "/" & cAfter
   
   Set sh = RecordSetToWorkSheet(rs, wb, shName)
   
End Sub

Sub Test_MergeRecordset_03()
   
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long

   Dim rs, rs1, rs2, rs3, rs4 As Recordset
   
   Set wb = ActiveWorkbook
   shName = "MergeRecordSet TEST 03"

   
   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"), adVariant)
   Set rs2 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q2"), adVarChar, VARCHAR_SIZE)
   cBefore = rs1.RecordCount + rs2.RecordCount
   
   
   ' At present there is a problem when adding adVariant data to adVarChar columns,
   ' so specify the the adVariant rs first in the MergeRecordset()
   Set rs = MergeRecordset(rs1, rs2)
   cAfter = rs.RecordCount
   
   'MsgBox "Records before/after MergeRecordst() " & cBefore & "/" & cAfter
   
   Set sh = RecordSetToWorkSheet(rs, wb, shName)
   
End Sub


Sub Test_GroupRecordset_01()
   
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long

   Dim rs, rs1, rs2, rs3, rs4 As Recordset
   
   Set wb = ActiveWorkbook
   shName = "GroupRecordSet TEST 01"
   
   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"))
   Set rs2 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q2"))
   Set rs3 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q3"))
   Set rs4 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q4"))
   Set rs = MergeRecordset(rs1, rs2, rs3, rs4)
   
   cBefore = rs1.RecordCount + rs2.RecordCount + rs3.RecordCount + rs4.RecordCount
   
   Set rs = GroupRecordSet(rs, "Company", "January")
   cAfter = rs.RecordCount
   
   'ErrorChange "Records before/after GroupRecordSet() " & cBefore & "/" & cAfter
   
   Set sh = RecordSetToWorkSheet(rs, wb, shName)
   
End Sub

Sub Test_GroupRecordset_02()
   
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long
   Dim grpColumns() As String
   Dim valColumns() As String
   
   
   Dim rs, rs1, rs2, rs3, rs4 As Recordset
   
   Set wb = ActiveWorkbook
   shName = "GroupRecordSet TEST 02"
   
   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"))
   Set rs2 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q2"))
   Set rs3 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q3"))
   Set rs4 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q4"))
   Set rs = MergeRecordset(rs1, rs2, rs3, rs4)
   
   cBefore = rs1.RecordCount + rs2.RecordCount + rs3.RecordCount + rs4.RecordCount
   
   
   grpColumns = Split("Company,Company Name", ",")
   valColumns = Split("January,February,March,April,May,June,July,August,September,October,November,December", ",")
   
   Set rs = GroupRecordSet(rs, grpColumns, valColumns)
   cAfter = rs.RecordCount
   
   'ErrorChange "Records before/after GroupRecordSet() " & cBefore & "/" & cAfter
   
   Set sh = RecordSetToWorkSheet(rs, wb, shName)
   
End Sub


Sub Test_GroupRecordset_03()
   
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long
   Dim grpColumns() As String
   Dim valColumns() As String
   
   
   Dim rs, rs1, rs2, rs3, rs4 As Recordset
   
   Set wb = ActiveWorkbook
   shName = "GroupRecordSet TEST 03"

   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"))
   Set rs2 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q2"))
   Set rs3 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q3"))
   Set rs4 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q4"))
   Set rs = MergeRecordset(rs1, rs2, rs3, rs4)
   
   cBefore = rs1.RecordCount + rs2.RecordCount + rs3.RecordCount + rs4.RecordCount
   
   
   grpColumns = Split("Company,Company Name,Account Code,Account Description,Made up column", ",")
   valColumns = Split("January,February,March,April,May,June,July,August,September,October,November,December", ",")
   
   Set rs = GroupRecordSet(rs, grpColumns, valColumns)
   cAfter = rs.RecordCount
   
   'ErrorChange "Records before/after GroupRecordSet() " & cBefore & "/" & cAfter
   
   Set sh = RecordSetToWorkSheet(rs, wb, shName)
   
End Sub

Sub Test_GroupRecordset_04()
   
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long
   Dim grpColumns() As String
   Dim valColumns() As String
   
   
   Dim rs, rs1, rs2, rs3, rs4, rs5 As Recordset
   
   Set wb = ActiveWorkbook
   shName = "GroupRecordSet TEST 04"
   
   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"))
   Set rs2 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q2"))
   Set rs3 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q3"))
   Set rs4 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q4"))
   Set rs5 = WorkSheetToRecordSet(wb.Worksheets("Budgets"))
   Set rs = MergeRecordset(rs1, rs2, rs3, rs4, rs5)
   
   cBefore = rs1.RecordCount + rs2.RecordCount + rs3.RecordCount + rs4.RecordCount + rs5.RecordCount
   
   
   grpColumns = Split("Company,Company Name", ",")
   valColumns = Split("January,February,March,April,May,June,July,August,September,October,November,December," & _
                      "January Budget,February Budget,March Budget,April Budget,May Budget,June Budget,July Budget,August Budget,September Budget,October Budget,November Budget,December Budget", _
                      ",")
   
   Set rs = GroupRecordSet(rs, grpColumns, valColumns)
   cAfter = rs.RecordCount
   
   'ErrorChange "Records before/after GroupRecordSet() " & cBefore & "/" & cAfter
   
   Set sh = RecordSetToWorkSheet(rs, wb, shName)
   
End Sub


Sub Test_PivotRecordset_01()
   
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long

   Dim rs, rs1 As Recordset
   
   Set wb = ActiveWorkbook
   shName = "PivotRecordSet TEST 01"
   
   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"), adVarChar, 1200)
   
   cBefore = rs1.RecordCount
   
   Set rs = PivotRecordSet(rs1, "Company Name", "Account Description", "January")
   cAfter = rs.RecordCount
   
   'ErrorChange "Records before/after GroupRecordSet() " & cBefore & "/" & cAfter
   
   Set sh = RecordSetToWorkSheet(rs, wb, shName)
   
End Sub


Sub Test_PivotRecordset_02()
   
   Dim sh As Worksheet
   Dim wb As Workbook
   Dim shName As String
   Dim cBefore As Long
   Dim cAfter As Long

   Dim rs, rs1 As Recordset
   
   Set wb = ActiveWorkbook
   shName = "PivotRecordSet TEST 02"
   
   Set rs1 = WorkSheetToRecordSet(wb.Worksheets("Actuals Q1"))
   
   cBefore = rs1.RecordCount
   
   Set rs = PivotRecordSet(rs1, "Company Name", "Account Description", "January")
   cAfter = rs.RecordCount
   
   'ErrorChange "Records before/after GroupRecordSet() " & cBefore & "/" & cAfter
   
   Set sh = RecordSetToWorkSheet(rs, wb, shName)
   
End Sub
