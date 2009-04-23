Attribute VB_Name = "SharedUtilitiesCSV"
Option Explicit

Function CSVToRecordSet(csvText As String, Optional datesAsUniversal As Boolean = False, Optional fieldDelimiter As String = ",", Optional fieldType As DataTypeEnum = adVariant, Optional fieldSize As ADO_LONGPTR, Optional fieldAttrib As FieldAttributeEnum) As ADODB.Recordset
  ' This doesnt handle quotes at all, it's a straight split on commas
  ' duplicate field names are ignored and the second value used
  Dim lines() As String
  Dim line As String
  Dim lineDelimiter As String
  Dim flds() As String
  Dim fldName As String
  Dim str As String
  
  
  Dim m As Long
  Dim i As Long
  Dim c As Long
  Dim eol As Long
  
  Dim rs As ADODB.Recordset

  
  If Len(csvText) < 1 Then
    Set CSVToRecordSet = Nothing
    Exit Function
  End If
  
  
  
  
  lineDelimiter = vbCrLf ' \r\n
  eol = InStr(csvText, lineDelimiter)
  If eol < 1 Then
    lineDelimiter = Chr(&HA) ' just \n
  End If
  
  lines = Split(csvText, lineDelimiter)
  

  
  m = UBound(lines)
  i = 0
  line = lines(i)
  
  Set rs = New ADODB.Recordset
  
  ' first line defines the field names
  flds = Split(line, fieldDelimiter)
  Do While c <= UBound(flds)
    fldName = Trim(flds(c))
    If fldName = "" Then
      fldName = "Field " & Lpad(CStr(c), "0", 3)
    End If
    
    If Not InFields(rs, fldName) Then
      rs.Fields.append fldName, fieldType, fieldSize, fieldAttrib
    End If
    c = c + 1
  Loop
  rs.Open
  
  ' Second line on
  i = i + 1
  Do While i <= m
    line = lines(i)
    If Len(line) > 0 Then
      flds = Split(line, fieldDelimiter)
      
      rs.AddNew
      c = 0
      Do While c <= UBound(flds)
        str = flds(c)
        
        If fieldType <> adVarChar Then
          rs.Fields(c).Value = CVarFromStr(str, datesAsUniversal)
        Else
          rs.Fields(c).Value = str
        End If
        
        c = c + 1
      Loop
    End If
    
    i = i + 1
  Loop
  
  rs.MoveFirst
  Set CSVToRecordSet = rs
  
End Function



