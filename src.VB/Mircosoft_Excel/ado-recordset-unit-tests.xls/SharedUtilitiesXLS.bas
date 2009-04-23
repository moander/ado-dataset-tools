Attribute VB_Name = "SharedUtilitiesXLS"
Option Explicit
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' TODO: Find solution for when there is over 65536 - 1 rows in a recordset


Public Function WorkSheetToRecordset(sh As Worksheet, Optional fieldType As DataTypeEnum = adVariant, Optional fieldSize As ADO_LONGPTR, Optional fieldAttrib As FieldAttributeEnum) As ADODB.Recordset
  ' fieldType has been tested with adVariant and adVarChar (you'll need a size)
  ' Any other value would be a bit strange
  Dim rs As ADODB.Recordset

  Dim f As Variant
  Dim cols As Long
  Dim rows As Long
  Dim c As Long
  Dim r As Long

  Dim ref As String
  Dim fldName As String
  
  ref = FindLastCell(sh)
  cols = sh.Range(ref).Column
  rows = sh.Range(ref).row
  
  Set rs = New ADODB.Recordset
  
  If ref <> "$A$1" Or sh.Range(ref).Value <> "" Then ''# This is to catch empty sheet
      c = 1
      r = 1
      Do While c <= cols
        fldName = sh.Cells(r, c).Value
        
        If InFields(rs, fldName) Then
          fldName = fldName + "01"
        End If
        
        rs.Fields.append fldName, fieldType, fieldSize, fieldAttrib
        c = c + 1
      Loop
      rs.Open


      r = 2
      Do While r <= rows
        rs.AddNew
        c = 1
        Do While c <= cols
          rs.Fields(c - 1) = sh.Cells(r, c).Value
          c = c + 1
        Loop
        r = r + 1
        'Debug.Print sh.name & ": " & r & " of " & rows & ", " & c & " of " & cols
      Loop

    End If
    
    Set WorkSheetToRecordset = rs
End Function



Public Function RecordsetToWorkSheet(ByVal rs As ADODB.Recordset, wb As Workbook, Optional wsName As String, Optional customNumberFormat As String = "mmm 'yy") As Worksheet
  Dim sh As Worksheet
  Dim v As Variant
  
  Dim f As Field
  Dim c As Long
  Dim r As Long
  Dim ref As String
  Dim prevState As Boolean
  
  If Trim(wsName) <> "" Then
    If InWorkBook(wb, wsName) Then
      prevState = Application.DisplayAlerts
      Application.DisplayAlerts = False
      wb.Worksheets(wsName).Delete
      Application.DisplayAlerts = prevState
    End If
    
    Set sh = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
    sh.name = wsName
  Else
    Set sh = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
    wsName = sh.name
  End If
  
  If rs Is Nothing Then
    sh.Activate
    Set RecordsetToWorkSheet = sh
    Exit Function
  End If

  rs.MoveFirst
  r = 1
  c = 1
  For Each f In rs.Fields
    sh.Cells(r, c).Value = CVarFromStr(f.name)
    c = c + 1
  Next
  rs.MoveFirst
  
  ' Using CopyFromRecordset() is a lot faster than looping through the recordset....
  ' however the type is very important
  ' NB: Tested for 8 columns and 99,472 rows in Excel 2007... but it only returns 65536 - 1 (for the header) rows.
  '     This is probably the Excel limit
  If rs.RecordCount > 0 Then
    rs.MoveLast ' dumb internet suggestion to combat random and very occasional errors with CopyFromRecordset()
    rs.MoveFirst
    StatusChange ("CopyFromRecordset() for " & rs.RecordCount & " rows.")
    ' Since adding support for adVariant, i've found a lot of errors appearing with CopyFromRecordset
    ' generally this can be fixed in the SharedRecordSet.bas by using .value for all fields eg:
    '   r.Fields("My Column").value = t.Fields("My Column").value
    ' as opposed to reliing on the default
    '   r.Fields("My Column") = t.Fields("My Column")
    
    sh.Range("A2").CopyFromRecordset rs, 65536 ' TransposeDim() is suggested as solution to large RS errors: http://support.microsoft.com/kb/q246335/
    StatusChange ("Finished CopyFromRecordset() for " & rs.RecordCount & " rows.")
  End If

  
  
  
  ' FORMATTING here down... probably should separate out but it's only for testing.
  ' Format the headers
  sh.rows(1).Font.Bold = True
  sh.rows(1).HorizontalAlignment = xlHAlignCenter
  sh.rows(1).VerticalAlignment = xlVAlignTop
  sh.rows(1).RowHeight = 35
  sh.rows(1).WrapText = True
  If customNumberFormat <> "" Then
    sh.rows(1).NumberFormat = customNumberFormat
  End If
  
  'sh.Columns.ColumnWidth = 30
  sh.Columns.AutoFit
  
  sh.Range("A2").Select
  
  ' You need to do all this to Freeze panes
  sh.Activate
  ActiveWindow.FreezePanes = True
  
  
  Set RecordsetToWorkSheet = sh
End Function

Public Function FindLastCell(sh As Worksheet) As String
  ' See http://www.ozgrid.com/VBA/ExcelRanges.htm for more examples
  ' This does not find empty cells with formatting in them
  Dim lastCol As Integer
  Dim lastRow As Long
  Dim lastCell As Range

  Dim Cells As Range
  Set Cells = sh.Cells
  
  If WorksheetFunction.CountA(Cells) > 0 Then
    lastRow = Cells.Find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row 'Search backwards by Rows.
    lastCol = Cells.Find(What:="*", After:=[a1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column 'Search backwards by Columns.
  End If
  
  If lastRow = 0 And lastCol = 0 Then
    FindLastCell = "A1"
  Else
    FindLastCell = Cells(lastRow, lastCol).Address
  End If
  
End Function

Public Function FindEndCell(sh As Worksheet) As String
  ' FindLastCell() is better
  Dim cols As Long
  Dim rows As Long
  Dim maxCols As Long
  Dim maxRows As Long
  Dim c As Long
  Dim r As Long

  maxRows = sh.rows.Count
  maxCols = sh.Columns.Count

  cols = sh.Range("A1").End(xlToRight).Column
  If cols >= maxCols Then
      cols = 1
  End If


  c = 1
  Do While c <= cols

    r = sh.Cells(1, c).End(xlDown).row
    If r >= maxRows Then
      r = 1
    End If

    If r > rows Then
      rows = r
    End If
    c = c + 1
  Loop

  FindEndCell = sh.Cells(rows, cols).Address

End Function

Public Function InWorkBook(wb As Workbook, wsName As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long
  
  InWorkBook = False
  Set var = Nothing
  
  Err.Clear
  On Error Resume Next
    Set var = wb.Worksheets(wsName)
    errNumber = CLng(Err.Number)
  On Error GoTo 0
  
 
  'MsgBox errNumber & " - " & Err.Number & " " & Err.Description
  
  If errNumber = 9 Then ' it's 9 if not in workbook
    InWorkBook = False
  ElseIf errNumber = 0 Then ' Or errNumber = 438 Then ' for some reason 438...
    InWorkBook = True
  Else
    MsgBox "InWorkBook() Error : " & errNumber & " - " & Err.Number & " " & Err.Description
  End If
  
  

  
End Function
