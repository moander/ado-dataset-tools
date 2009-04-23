Attribute VB_Name = "SharedRecordSet"
' This module is the property of Mark Nold
' It is available for use via the LGPL license.


' Requires reference to "Microsoft ActiveX Data Object Library 2.7" or above

Option Explicit

Public Const VARCHAR_SIZE = 1200 ' This one is important. I use 1200 since Ellipse COM components can create data up to this size (MSO010 extended text)

Public Const ADT_ACTION_SUM As String = "sum"
Public Const ADT_ACTION_MAX As String = "max"
Public Const ADT_ACTION_MIN As String = "min"
Public Const ADT_ACTION_CONCATENATE As String = "concat"
Public Const ADT_ACTION_JOIN_COMMA As String = "joinc"
Public Const ADT_ACTION_JOIN_HASH As String = "joinh"
Public Const ADT_ACTION_JOIN_CRLF As String = "joinh"





Enum SortRecordSetOptionsEnum
    TEXT_AS_NUMERIC = 1
End Enum

Enum GroupRecordSetOptionsEnum
    SORT_BY_GC = 1
    NULLS_TO_ZERO = 2
    ZEROS_TO_NULL = 4
    'SORT_BY_PC = 6 ' Sort by Pivot Column (useful for dates) : NOT IMPLEMENTED
End Enum


Enum SubTotalRecordSetOptionsEnum
    ST_SORT_BY_GC = 1
    ST_NULLS_TO_ZERO = 2
    ST_ZEROS_TO_NULL = 4
    'ST_SORT_BY_PC = 6 ' Sort by Pivot Column (useful for dates) : NOT IMPLEMENTED

    GRAND_TOTAL = 8

    'ADD_TOTAL_COLUMN = 64
    BLANK_AFTER = 128
    NO_SINGLE_TOTALS = 256
    TOTAL_TEXT_INCLUDE_DATA = 512

End Enum

Enum GroupRecordsetAggregateEnum
 ' NB: Enums to replace "actions" on GroupRecordset() etc
 '     not implemented becuase currently GroupRecordset() happily accepts an array of actions
 '     or a single string value, otherwise defaults to "sum"
  [_First] = 1
  adtSum = 1
  adtMax = 2
  adtMin = 3
  adtConcatenate = 4
  
  ' These should be one option and a string passed as the join / concatenate string
  adtJoinc = 5 ' Join with Comma
  adtJoinh = 6 ' Join with Hash
  adtJoinCrlf = 7 ' Join with crlf
  [_Last] = 7
End Enum


Function FieldExists(r As ADODB.Recordset, fieldName As String) As Boolean
  ' This is depricated in favour of InFields() .. in fact the function is the same
  Dim foundFieldName As String
  FieldExists = False
  
  
  On Error Resume Next
    foundFieldName = r.Fields(fieldName).name
  On Error GoTo 0

  If foundFieldName = fieldName Then
    FieldExists = True
  End If
 
End Function

Function InFields(r As ADODB.Recordset, fieldName As String) As Boolean
  Dim foundFieldName As String
  InFields = False

  On Error Resume Next
    foundFieldName = r.Fields(fieldName).name
  On Error GoTo 0

  If foundFieldName = fieldName Then
    InFields = True
  End If
 
End Function

Function FindItem(clx As Collection, key As String) As Variant ' NB: Variant disappears in VB.net
  ' it would be nice to simply extend the collection object but not so easy,
  On Error GoTo NotFound
    FindItem = clx.Item(key)
  Exit Function
  
NotFound:
  Set FindItem = Nothing
  Exit Function

End Function


Function CopyRecordIntoRecordset(Recordset As ADODB.Recordset, row As ADODB.Fields) As ADODB.Recordset
  ' this needs some error checking
  Dim f As Variant
  Dim fieldName As String
  Dim fieldValue As String
  
  ' Do we need some new fields in the recordset defn?
  If Recordset Is Nothing Then
    Set Recordset = New ADODB.Recordset
    Recordset.CursorType = adOpenKeyset
    Recordset.LockType = adLockOptimistic
  End If
  
  ' Really should check to see if it is open or not...
  If Recordset.Fields.Count < 1 Then
    For Each f In row
      Recordset.Fields.append CStr(f.name), f.Type, f.DefinedSize

    Next
    Recordset.Open
  End If
  
  
  Recordset.AddNew
  For Each f In row
    fieldName = f.name
    fieldValue = f.Value
    ' should check if fieldName exists
    Recordset.Fields(fieldName) = f.Value
  Next
  Recordset.Update
  
  Set CopyRecordIntoRecordset = Recordset
End Function


Function CloneEmptyRecordset(r As ADODB.Recordset) As ADODB.Recordset
  ' This should be merged into CopyRecordsetStructure()

  ' This is just a cheap and nasty way of cloning a recordset then deleting everything in it
  ' i would assume s.Delete adAffectGroup would improve this but probably should build a better
  ' CloneRecordset()
  
  ' Currently CloneRecordsetStructure() is a better way of doing this.
  
  Dim n As ADODB.Recordset
  
  If r Is Nothing Then
    Set CloneEmptyRecordset = Nothing
    Exit Function
  End If


    
  Set n = CloneRecordset(r.Clone)
  ' i though i could use
  ' s.Delete adAffectGroup
  ' but it appears not.
  
  'MsgBox "BEFORE " & n.RecordCount
  Do Until n.EOF Or n.RecordCount = 0
    n.MoveFirst
    n.Delete
  Loop
  
  n.UpdateBatch
  
  Set CloneEmptyRecordset = n
End Function


Function CloneRecordsetStructure(r As ADODB.Recordset, Optional OpenOnCreate As Boolean = True) As ADODB.Recordset
  ' This should be merged into CopyRecordsetStructure()
    Dim fld As ADODB.Field
    Dim oRsCloned As ADODB.Recordset
    
    Set oRsCloned = New ADODB.Recordset
    
    For Each fld In r.Fields
        oRsCloned.Fields.append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
        
        'special handling for data types with numeric scale & precision
        Select Case fld.Type
            Case adNumeric, adDecimal
                oRsCloned.Fields(oRsCloned.Fields.Count - 1).Precision = fld.Precision
                oRsCloned.Fields(oRsCloned.Fields.Count - 1).NumericScale = fld.NumericScale
        End Select
    Next
    
    'make the cloned recordset ready for business
    If OpenOnCreate = True Then
      oRsCloned.Open
    End If
    
    'return the new recordset
    Set CloneRecordsetStructure = oRsCloned
    
    'clean up
    Set fld = Nothing
End Function

Function CloneRecordsetStructureAsVariant(r As ADODB.Recordset, Optional OpenOnCreate As Boolean = True) As ADODB.Recordset
  ' This should be merged into CopyRecordsetStructure()
    Dim fld As ADODB.Field
    Dim t As ADODB.Recordset
    
    Set t = New ADODB.Recordset
    
    For Each fld In r.Fields
        t.Fields.append fld.name, adVariant
    Next
    
    'make the cloned recordset ready for business
    If OpenOnCreate = True Then
      t.Open
    End If
    
    'return the new recordset
    Set CloneRecordsetStructureAsVariant = t
    
    'clean up
    Set fld = Nothing
End Function

Function CloneRecordsetStructureAsVarChar(r As ADODB.Recordset, Optional OpenOnCreate As Boolean = True) As ADODB.Recordset
  ' This should be merged into CopyRecordsetStructure()
    Dim fld As ADODB.Field
    Dim t As ADODB.Recordset
    
    Set t = New ADODB.Recordset
    
    For Each fld In r.Fields
        t.Fields.append fld.name, adVarChar, VARCHAR_SIZE
    Next
    
    'make the cloned recordset ready for business
    If OpenOnCreate = True Then
      t.Open
    End If
    
    'return the new recordset
    Set CloneRecordsetStructureAsVarChar = t
    
    'clean up
    Set fld = Nothing
End Function
Function CloneRecordset(ByVal rs As ADODB.Recordset, Optional ByVal LockType As ADODB.LockTypeEnum = -1) As ADODB.Recordset
  ' See  http://www.vbrad.com/article.aspx?id=12 for a discussion on memory leaks with ADODB.Stream
  
  ' Also from http://www.vbrad.com/article.aspx?id=12
  ' Contrary to popular belief, Recordset.Clone doesn't actually clone the recordset.
  ' It doesn't actually create a new object in memory - it simply returns a reference
  ' to the original recordset with the option of making the reference read-only.
  ' To verify this claim, simply delete a record from the cloned recordset and
  ' you will see that the .RecordCount on the original recordset also decreases.

  
  Dim stm As ADODB.Stream
  Dim cln As ADODB.Recordset
  Dim f As Field
  Dim adVariantFlag As Boolean
  
  adVariantFlag = False
  
  If rs Is Nothing Then
    Set CloneRecordset = cln
    Exit Function
  End If
  
  
  
  If rs.Fields.Count < 1 Then
    Set CloneRecordset = cln
    Exit Function
  End If
  
  
  For Each f In rs.Fields
    If f.Type = adVariant Then
      adVariantFlag = True
    End If
  Next
  
  
  
  If adVariantFlag = False Then
    'save the recordset to the stream object
    Set stm = New ADODB.Stream
    rs.Save stm
    'and now open the stream object into a new recordset
    Set cln = New ADODB.Recordset
    cln.Open stm, , , LockType
  Else
    ' streams dont support adVariant.
    ' You could convert it all to text by
    ' rs.Save stm, adPersistXML
    ' but then why not just use adVarChar?
    Set cln = CopyRecordsetStructure(rs)
    Dim flds As Fields
    
    rs.MoveFirst
    Do Until rs.EOF
    
      ' cln.AddNew fldNames, fldValues
      cln.AddNew
      For Each f In rs.Fields
        cln.Fields(f.name) = f.Value
      Next
      
      rs.MoveNext
    Loop
  End If
  
  'return the cloned recordset
  Set CloneRecordset = cln
  
  'release the reference
  Set cln = Nothing
  Set stm = Nothing
 
End Function


Function CopyRecordsetStructure(r As ADODB.Recordset, Optional OpenOnCreate As Boolean = True) As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim cpy As ADODB.Recordset
    
    Set cpy = New ADODB.Recordset
    
    For Each fld In r.Fields
        cpy.Fields.append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
        
        'special handling for data types with numeric scale & precision
        Select Case fld.Type
            Case adNumeric, adDecimal
                cpy.Fields(cpy.Fields.Count - 1).Precision = fld.Precision
                cpy.Fields(cpy.Fields.Count - 1).NumericScale = fld.NumericScale
        End Select
    Next
    
    'make the cloned recordset ready for business
    If OpenOnCreate = True Then
      cpy.Open
    End If
    
    'return the new recordset
    Set CopyRecordsetStructure = cpy
    
    'clean up
    Set fld = Nothing
End Function

Function CloneRecordset_OLD(ByVal oRs As ADODB.Recordset, _
  Optional ByVal LockType As ADODB.LockTypeEnum = -1) As ADODB.Recordset

  Dim oStream As ADODB.Stream
  Dim oRsClone As ADODB.Recordset
  
  If oRs Is Nothing Then
    Set CloneRecordset = oRsClone
    Exit Function
  End If
  
  
  
  If oRs.Fields.Count < 1 Then
    Set CloneRecordset = oRsClone
    Exit Function
  End If
  
  
  
  'save the recordset to the stream object
  Set oStream = New ADODB.Stream
  oRs.Save oStream
    
  'and now open the stream object into a new recordset
  Set oRsClone = New ADODB.Recordset
  oRsClone.Open oStream, , , LockType
  
  'return the cloned recordset
  Set CloneRecordset = oRsClone
  
  'release the reference
  Set oRsClone = Nothing
  Set oStream = Nothing ' Does nothing though.
 
End Function


Function CreateVarCharRecordsetFromString(strFields As String, Optional OpenOnCreate As Boolean = True, Optional strIndex As String) As ADODB.Recordset
  
  Dim myFields() As String
  Dim myIndexes() As String
  Dim myVerified() As String
  Dim d As Variant
  Dim i As Long
  Dim fieldName As String

  
  Dim r As ADODB.Recordset
  Set r = New ADODB.Recordset
  
  myFields = Split(strFields, ",")
  myIndexes = Split(strIndex, ",")
  
  If IsArray(myIndexes) Then
    For i = 0 To UBound(myIndexes)
      If InArray(myFields, myIndexes(i)) Then
        ReDim Preserve myVerified(i) As String
        myVerified(i) = myIndexes(i)
      End If
    Next
    strIndex = Join(myVerified, ",")
    r.Index = strIndex
  Else
    strIndex = ""
  End If
  
   
  
  For Each d In myFields
    fieldName = CStr(d)
    If Len(Trim(fieldName)) > 0 Then
      StatusChange "Adding field : " & fieldName
      r.Fields.append fieldName, adVarChar, VARCHAR_SIZE  ' This is an Ellipse Connector hangover. Probably should type the data but stay with text for interoperability. Unfortuntaly you cant use adVarient if you want to sort and StandardText can be up to 20 lines of 60 characters.
    Else
      ErrorChange "Attempting to add a blank field"
    End If
  Next
  StatusChange r.Fields.Count & " Fields in r"
  
  If OpenOnCreate = True Then
    r.Open
  End If



  Set CreateVarCharRecordsetFromString = r
  
  
End Function


Function PivotRecordSet(ByRef r As ADODB.Recordset, groupColumns As Variant, pivotColumns As Variant, valueColumn As String, Optional action As String, Optional ByVal options As GroupRecordSetOptionsEnum = 0) As ADODB.Recordset
  StatusChange "PivotRecordSet() starting"
  ' Set rt = PivotRecordSet(r, Split("ResourceGroup,ResourceID,ResourceName,ResourceType,DailyRate", ","), "ReportedWeekEnding", "DaysWorked", "sum")
  ' The order of columns seem to depend on the original order and
  ' not the specified GroupColumns
  
  
  Dim t As ADODB.Recordset
  Dim x As ADODB.Recordset
  Dim h As New HashTable
  
  
  Dim gc() As String
  Dim gc2() As String
  Dim pc() As String
  
  
  Dim cols() As String
  ReDim cols(0)
  Dim uniqueValueColumns() As String
  ReDim uniqueValueColumns(0)
  
  Dim i, l, u As Long
  Dim c As Variant
  Dim strName As String
  Dim strFilter As String
  Dim prvCol As String
  Dim fld As Variant
  
  Dim valType As DataTypeEnum
  Dim valSize As ADO_LONGPTR
  Dim valAttributes As FieldAttributeEnum
  Dim valPrecision As Variant ' no idea of the type
  Dim valNumericScale As Variant ' no idea of the type
  
  
  
  
  
  gc = CStrArray(groupColumns)
  gc2 = CStrArray(groupColumns)
  pc = CStrArray(pivotColumns)
  
  ' Initially we can handle only 1 pivotColumn
  ' gc2() is used just to pass to the GroupRecordSet function
  ReDim Preserve gc2(UBound(gc2) + 1)
  gc2(UBound(gc2)) = pc(0)
  
  Set t = GroupRecordSet(r, gc2, valueColumn, action, options)
  Set x = New ADODB.Recordset
  
  
  
  
  
  
  For Each fld In t.Fields
    If InArray(pc, fld.name) = False And fld.name <> valueColumn Then
      'x.Fields.Append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
      
      'special handling for data types with numeric scale & precision
      'Select Case fld.Type
      '    Case adNumeric, adDecimal
      '        x.Fields(x.Fields.Count - 1).Precision = fld.Precision
      '        x.Fields(x.Fields.Count - 1).NumericScale = fld.NumericScale
      'End Select
      Select Case fld.Type
      Case adVariant
        x.Fields.append fld.name, fld.Type
      Case adNumeric, adDecimal
        x.Fields.append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
        x.Fields(x.Fields.Count - 1).Precision = fld.Precision
        x.Fields(x.Fields.Count - 1).NumericScale = fld.NumericScale
      Case Else
        x.Fields.append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
      End Select
      
      
    End If
  Next


  valType = t.Fields(valueColumn).Type
  valSize = t.Fields(valueColumn).DefinedSize
  valAttributes = t.Fields(valueColumn).Attributes
  
  valPrecision = t.Fields(valueColumn).Precision
  valNumericScale = t.Fields(valueColumn).NumericScale
  
  t.MoveFirst
  Do Until t.EOF
    'strName = CStr(t.Fields(pc(0)).Value)
    strName = CStrFromVar(t.Fields(pc(0)).Value)

    If strName <> prvCol Then ' minor saving if sorted
      ReDim Preserve cols(UBound(cols) + 1)
      cols(UBound(cols)) = strName
    End If
    
    prvCol = strName
    t.MoveNext
  Loop
 
  
  
  cols = SortStringArray(cols, True) ' THIS IS A BIG PROBLEM if COLS is dates or some such.
  i = 0
  For c = 0 To UBound(cols)
    strName = CStr(cols(c))
    
    StatusChange c & " - " & strName & " " & InArray(pc, strName) & " FROM: " & Join(pc, ",")
    
    If Len(Trim(strName)) > 0 And FieldExists(x, strName) = False Then
      StatusChange "Adding field : " & strName
      
      Select Case valType
      Case adVariant
        x.Fields.append strName, valType
      Case adNumeric, adDecimal
        x.Fields.append strName, valType, valSize, valAttributes
        x.Fields(x.Fields.Count - 1).Precision = valPrecision
        x.Fields(x.Fields.Count - 1).NumericScale = valNumericScale
      Case Else
        x.Fields.append strName, valType, valSize, valAttributes
      End Select
      
      ReDim Preserve uniqueValueColumns(i)
      uniqueValueColumns(i) = strName
      i = i + 1
      
    End If
  Next
  
  'x.Fields.Append "INFO", adVarChar, 1200
  x.Open
  
 
  l = UBound(gc)
  u = UBound(pc)
  Dim strTmp As String

  t.MoveFirst
  Do Until t.EOF

    ' x.Find, x.Filter don't work with adVariant columns
    ' x.Seek and x.Index are not supported.
    ' x.Filter = BuildRSFilter(gc, t.Fields)


    strFilter = BuildRSFilter(gc, t.Fields)
    If h.Exists(strFilter) = False Then
      x.AddNew
      h.Add strFilter, x.RecordCount
      
      i = 0
      While i <= l
        strName = gc(i)
        x.Fields(strName).Value = t.Fields(strName).Value
        i = i + 1
      Wend
    Else
      x.Move (h.Item(strFilter) - x.AbsolutePosition)
    End If
    

    i = 0
    While i <= u
      'strName = t.Fields(pc(i)).Value
      strName = CStrFromVar(t.Fields(pc(i)).Value)
      'Debug.Print x.Fields(strName).Type & " " & adDate
      x.Fields(strName).Value = t.Fields(valueColumn).Value
      i = i + 1
    Wend
    
    'If x.RecordCount < 1 Then
    '  x.MoveFirst
    '  bMark = x.bookmark
    'End If
    t.MoveNext
  Loop
  x.Filter = ""
  StatusChange "PivotRecordSet " & x.RecordCount & " rows " & x.Fields.Count & " columns"
  
  
  Set PivotRecordSet = CloneRecordset(x)
  
End Function

Function SortRecordSet(r As ADODB.Recordset, sortColumn As String, Optional options As SortRecordSetOptionsEnum = 0) As ADODB.Recordset
  Dim t As ADODB.Recordset
  Dim x As ADODB.Recordset
  Dim s As ADODB.Recordset
  
  Dim max As Long
  Dim l As Long
  Dim c As Long
  Dim lngRowID As Long
  Dim f As Variant
  
  If InFields(r, sortColumn) = False Then
    Set SortRecordSet = CloneRecordset(r)
    Exit Function
  End If
  
  
  Set t = CloneRecordset(r)
  If options = 0 Then
    
    t.Sort = sortColumn
    Set SortRecordSet = CloneRecordset(t)
    Exit Function
  End If
  

  c = 1

  t.MoveFirst
  Do Until t.EOF
    l = Len(t.Fields(sortColumn))
    If l > max Then
      max = l
    End If
    t.MoveNext
  Loop


  Set x = New ADODB.Recordset
  x.Fields.append "RowID", adSingle
  x.Fields.append "OriginalValue", adVarChar, max
  x.Fields.append "NumericValue", adVarChar, max + 10
  x.Open
  

  t.MoveFirst
  Do Until t.EOF
    x.AddNew
    x.Fields("RowID") = c
    x.Fields("OriginalValue") = t.Fields(sortColumn)
    x.Fields("NumericValue") = Lpad(t.Fields(sortColumn), "0", max)
    
    t.MoveNext
    c = c + 1
  Loop
  x.Sort = "NumericValue"
  
  Set s = CloneRecordsetStructure(t)
  
  x.MoveFirst
  t.MoveFirst
  Do Until x.EOF
    lngRowID = x.Fields("RowID")
    t.Move (lngRowID - t.AbsolutePosition)
    s.AddNew
    For Each f In t.Fields
      s.Fields(f.name) = f.Value
    Next
    
    x.MoveNext
  Loop

  Set SortRecordSet = s
End Function

Function DeleteFields(r As ADODB.Recordset, deleteColumns As Variant) As ADODB.Recordset
  Dim t As ADODB.Recordset
  Dim s As ADODB.Recordset
  
  Dim dc() As String
  
  Dim f As Variant
  Dim strName As String
  
  dc = CStrArray(deleteColumns)
  Set t = CloneRecordset(r)
  Set s = New ADODB.Recordset
  
  For Each f In t.Fields
    If InArray(dc, f.name) = False Then
      s.Fields.append f.name, f.Type, f.DefinedSize, f.Attributes
      
      'special handling for data types with numeric scale & precision
      Select Case f.Type
          Case adNumeric, adDecimal
              s.Fields(s.Fields.Count - 1).Precision = f.Precision
              s.Fields(s.Fields.Count - 1).NumericScale = f.NumericScale
      End Select
    End If
  Next
  s.Open
  
  
  
  t.MoveFirst
  Do Until t.EOF
    s.AddNew
    For Each f In s.Fields
      strName = f.name
      s.Fields(strName).Value = t.Fields(strName).Value
    Next
    t.MoveNext
  Loop
  
  Set DeleteFields = s
  
End Function

Function AddFields(ByVal r As ADODB.Recordset, addColumns As Variant, Optional defaultData As Variant = Nothing, Optional append = True, Optional fieldType As DataTypeEnum = adVariant, Optional fieldSize As ADO_LONGPTR, Optional fieldAttrib As FieldAttributeEnum) As ADODB.Recordset
  Dim t As ADODB.Recordset
  Dim s As ADODB.Recordset
  Dim ac() As String
  
  Dim f As Variant
  Dim strName As String
  
  Dim m As Long
  Dim i As Long
  
  ac = CStrArray(addColumns)
  Set t = CloneRecordset(r)
  Set s = New ADODB.Recordset
  
  m = UBound(ac)
  i = 0
  
  
  If append = False Then
    Set s = AddFieldsHelper(s, ac, fieldType, fieldSize, fieldAttrib)
    
  End If
  
  For Each f In t.Fields
    s.Fields.append f.name, f.Type, f.DefinedSize, f.Attributes
    'special handling for data types with numeric scale & precision
    Select Case f.Type
        Case adNumeric, adDecimal
            s.Fields(s.Fields.Count - 1).Precision = f.Precision
            s.Fields(s.Fields.Count - 1).NumericScale = f.NumericScale
    End Select
  Next
  
  If append = True Then
    Set s = AddFieldsHelper(s, ac, fieldType, fieldSize, fieldAttrib)
  End If
  s.Open
  
  
  
  Dim dFlag As Boolean

  If IsObject(defaultData) Then
    If Not defaultData Is Nothing Then
      dFlag = True
    End If
  Else
    If CStr(defaultData) <> "" Then
      dFlag = True
    End If
  End If
    
  

  
  t.MoveFirst
  Do Until t.EOF
    s.AddNew
    

    
    For Each f In s.Fields
      strName = f.name
      If InFields(t, strName) Then
        s.Fields(strName).Value = t.Fields(strName).Value
      ElseIf dFlag Then
        s.Fields(strName).Value = defaultData
      End If
    Next
    
    t.MoveNext
  Loop
  
  Set AddFields = s
  
End Function


Private Function AddFieldsHelper(rs As ADODB.Recordset, str() As String, Optional fieldType As DataTypeEnum = adVariant, Optional fieldSize As ADO_LONGPTR, Optional fieldAttrib As FieldAttributeEnum) As ADODB.Recordset
  Dim i As Long
  Dim m As Long
  
  m = UBound(str)
  i = 0
  Do While i <= m
     rs.Fields.append str(i), fieldType, fieldSize, fieldAttrib
     i = i + 1
  Loop
  Set AddFieldsHelper = rs
End Function
Function GroupRecordSet(ByVal r As ADODB.Recordset, groupColumns As Variant, valueColumns As Variant, Optional actions As Variant, Optional options As GroupRecordSetOptionsEnum = 0) As ADODB.Recordset
  ' groupColumns, valueColumns, actions are all assumed to be either a String or an array of strings
  
  ' if action is specified as "max" and there are more than one valueColumn then the "max" action will be used for all values
  
  ' Need to add the ability to undertake different actions on the same column
  ' eg: "DaysWorked",split("sum","max",",")
  '     This should result in two columns one called [DaysWorked SUM] and [DaysWorked MAX]
  
  ' Need to look into improving the speed without using SORT
  ' the order must always remain the same and it's up to the user to
  ' sort if they wish.
  
  'timing
  Dim startTime As Single
  Dim t1, t2 As Single
  
  
  StatusChange "GroupRecordSet() starting"
  
  Dim t As ADODB.Recordset
  Dim s As ADODB.Recordset
  Dim x As ADODB.Recordset
  Dim h As New HashTable
  
  Dim fld As Variant
  
  
  Dim strFilter As String
  Dim prevFilter As String
  Dim strName As String
  Dim strAction As String
  Dim aryFilter() As String
  ReDim aryFilter(0) As String
  
  Dim prvAction As String
  Dim prvSortID As Long
  Dim lngRowID As Long
  
  
  Dim gc() As String
  Dim vc() As String
  Dim ac() As String
  
  Dim gcTemp() As String
  Dim vcTemp() As String
  Dim acTemp() As String
  
  Dim i, c, u, f, m As Long
  Dim ubGroupColumns As Long ' NB: if this is -1 then no real group columns exist.
  ubGroupColumns = -1
  
  gc = CStrArray(groupColumns)
  vc = CStrArray(valueColumns)
  ac = CStrArray(actions)
  
  
  
  ' You cant have more actions than value columns
  m = UBound(vc)
  ReDim Preserve ac(m)
  
  prvAction = "sum"
  
  While i <= m
    ac(i) = LCase(ac(i))
    
    If ac(i) <> ADT_ACTION_SUM And _
       ac(i) <> ADT_ACTION_MIN And _
       ac(i) <> ADT_ACTION_MAX And _
       ac(i) <> ADT_ACTION_CONCATENATE And _
       ac(i) <> ADT_ACTION_JOIN_COMMA And _
       ac(i) <> ADT_ACTION_JOIN_HASH And _
       ac(i) <> ADT_ACTION_JOIN_CRLF Then

       
      ac(i) = prvAction
    Else
      prvAction = ac(i)
    End If
    
    i = i + 1
  Wend



  ' Check to make sure that all of gc and vc exist.
  



  Set t = CloneRecordset(r)
  Set s = New ADODB.Recordset
  
  m = UBound(gc)
  i = 0
  c = 0
  While i <= m
    strName = gc(i)
    If FieldExists(t, strName) Then
      
      If FieldExists(s, strName) = False Then
        Set fld = t.Fields(strName)
        s.Fields.append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
    
        'special handling for data types with numeric scale & precision
        Select Case fld.Type
            Case adNumeric, adDecimal
                s.Fields(s.Fields.Count - 1).Precision = fld.Precision
                s.Fields(s.Fields.Count - 1).NumericScale = fld.NumericScale
        End Select
      End If
      ReDim Preserve gcTemp(c) As String
      gcTemp(c) = gc(i)
      c = c + 1
    End If
    i = i + 1
  Wend
  
  m = UBound(vc)
  i = 0
  c = 0
  While i <= m
    strName = vc(i)
    If FieldExists(t, strName) Then
      If FieldExists(s, strName) = False Then
        Set fld = t.Fields(strName)
        s.Fields.append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
    
        'special handling for data types with numeric scale & precision
        Select Case fld.Type
            Case adNumeric, adDecimal
                s.Fields(s.Fields.Count - 1).Precision = fld.Precision
                s.Fields(s.Fields.Count - 1).NumericScale = fld.NumericScale
        End Select
      End If
      
      ReDim Preserve vcTemp(c) As String
      vcTemp(c) = vc(i)
      c = c + 1
    End If
    i = i + 1
  Wend
  
  
  gc = gcTemp
  vc = vcTemp
  If IsArray(gc, True) Then
    ubGroupColumns = UBound(gc)
  End If
  
  
  'make the recordset ready for business
  s.Open

  ' if SORT_BY_GC is set then sort by the group columns
  
  If (options And SORT_BY_GC) > 0 Then
    t.Sort = BuildRSSort(gc, t)
    Set t = CloneRecordset(t)
  Else
   ' do nothing
  End If


  Set x = New ADODB.Recordset
  x.Fields.append "SortID", adSingle
  x.Fields.append "RowID", adSingle
  x.Open
  
  t1 = Timer
  ' Add the Group Columns gc() to the summary s recordset
  c = 1
  t.MoveFirst
  Do Until t.EOF
    If ubGroupColumns < 0 Then
      strFilter = "TOTAL"
    Else
      strFilter = BuildRSFilter(gc, t.Fields)
    End If
    
    If h.Exists(strFilter) = False Then
      h.Add strFilter, h.Count + 1
    End If
    
    x.AddNew
    x.Fields("SortID") = h.Item(strFilter)
    x.Fields("RowID") = c
    
    t.MoveNext
    c = c + 1
  Loop
  x.Sort = "[SortID] , [RowID] ASC"
  'Set GroupRecordSet2 = CloneRecordset(x)
  'Exit Function
  
  t1 = Timer - t1
  t2 = Timer
  
  ' Add the Values vc() columns according the the actions ac()
  u = UBound(vc)
  
  prvSortID = 0
  x.MoveFirst
  t.MoveFirst
  Do Until x.EOF
    lngRowID = x.Fields("RowID")
    t.Move (lngRowID - t.AbsolutePosition)
    If x.Fields("SortID") <> prvSortID Then
      s.AddNew
      i = 0
      While i <= ubGroupColumns
        strName = gc(i)
        s.Fields(strName).Value = t.Fields(strName).Value
        i = i + 1
      Wend
      i = 0
      While i <= u
        strName = vc(i)
        s.Fields(strName).Value = t.Fields(strName).Value
        i = i + 1
      Wend
    Else

      i = 0
      While i <= u
        strName = vc(i)
        strAction = ac(i)
        ' NOTE: somefield.value = someotherfield.value is extremly important once
        '       adVariant columns are used.
        '
        'fldt = t.Fields(strName) ' not used YET
        'flds = t.Fields(strName) ' not used YET
        
        
        Select Case strAction
        Case ADT_ACTION_SUM
          If IsNumeric(t.Fields(strName).Value) = True Then
            If IsNumeric(s.Fields(strName).Value) = True Then
              s.Fields(strName).Value = CDbl(s.Fields(strName).Value) + CDbl(t.Fields(strName).Value)
            Else
              s.Fields(strName).Value = CDbl(t.Fields(strName).Value)
            End If
            
          End If
        Case ADT_ACTION_MIN
          If IsNumeric(t.Fields(strName).Value) = True Then
            If CDbl(t.Fields(strName).Value) <= CDbl(s.Fields(strName).Value) Then
              s.Fields(strName).Value = CDbl(t.Fields(strName).Value)
            End If
          End If
        Case ADT_ACTION_MAX
          If IsNumeric(t.Fields(strName).Value) = True Then
            If CDbl(t.Fields(strName).Value) >= CDbl(s.Fields(strName).Value) Then
              s.Fields(strName).Value = CDbl(t.Fields(strName).Value)
            End If
          End If
        Case ADT_ACTION_CONCATENATE
          s.Fields(strName).Value = CStr(s.Fields(strName).Value) & CStr(t.Fields(strName).Value)
        Case ADT_ACTION_JOIN_COMMA
          s.Fields(strName).Value = CStr(s.Fields(strName).Value) & "," & CStr(t.Fields(strName).Value)
        Case ADT_ACTION_JOIN_HASH
          s.Fields(strName).Value = CStr(s.Fields(strName).Value) & "#" & CStr(t.Fields(strName).Value)
        Case ADT_ACTION_JOIN_CRLF
          s.Fields(strName).Value = CStr(s.Fields(strName).Value) & vbCrLf & CStr(t.Fields(strName).Value)
        
        Case Else
          s.Fields(strName) = 0
        End Select
        i = i + 1
      Wend

    End If
    
    prvSortID = x.Fields("SortID")
    x.MoveNext
  Loop

  
  t2 = Timer - t2
  
  'MsgBox Format(t1, "#0.0000") & " vs " & Format(t2, "#0.0000")
  StatusChange "GroupRecordSet " & s.RecordCount & " rows " & s.Fields.Count & " columns"
  Set GroupRecordSet = CloneRecordset(s)
  
End Function

Function SubTotalRecordSet(r As ADODB.Recordset, groupColumns As Variant, valueColumns As Variant, Optional actions As Variant, Optional headerText As String = "", Optional totalText As String = "", Optional options As SubTotalRecordSetOptionsEnum = 0, Optional ByRef trackerTable As HashTable) As ADODB.Recordset
  ' groupColumns, valueColumns, actions are all assumed to be either a String or an array of strings
  
  ' if action is specified as "max" and there are more than one valueColumn then the "max" action will be used for all values
  
  ' Need to add the ability to undertake different actions on the same column
  ' eg: "DaysWorked",split("sum","max",",")
  '     This should result in two columns one called [DaysWorked SUM] and [DaysWorked MAX]
  
  ' Need to look into improving the speed without using SORT
  ' the order must always remain the same and it's up to the user to
  ' sort if they wish.
  
  'timing
  Dim startTime As Single
  Dim t1, t2 As Single
  
  '    ADD_TOTAL = 64
  '  BLANK_AFTER = 128
  '  BLANK_BEFORE = 256
    
  StatusChange "SubTotalRecordSet() starting"
  
  Dim t As ADODB.Recordset
  Dim g As ADODB.Recordset
  Dim s As ADODB.Recordset
  Dim x As ADODB.Recordset
  Dim h As New HashTable
  
  Dim fld As Variant
  
  Dim strFilter As String
  Dim prevFilter As String
  Dim strName As String
  Dim strAction As String
  Dim aryFilter() As String
  ReDim aryFilter(0) As String
  
  Dim aryTemp() As String
  
  Dim prvAction As String
  Dim prvSortID As Long
  Dim lngRowID As Long
  
  
  Dim gc() As String
  Dim gcTemp() As String
  
  Dim i, c, u, f, m As Long
  Dim ubGroupColumns As Long ' NB: if this is -1 then no real group columns exist.
  ubGroupColumns = -1
  
  gc = CStrArray(groupColumns)
  
  Set t = CloneRecordset(r)
  Set g = GroupRecordSet(r, groupColumns, valueColumns, actions, options)
  
  Set s = CloneRecordsetStructure(r, False)
  ' THIS NEEDS TO BE EXAMINED AND WHERE THE FUNCTIONALITY SHOULD LIVE
  'If (options And ADD_TOTAL_COLUMN) > 0 Then
  '  ' This should be the default type if there is confusion
  '  ' on the type and size.
  '  i = 1
  '  Do While i <= 99
  '    strName = "Total_" & Lpad(CStr(i), "0", 2)
  '    If InFields(s, strName) = False Then
  '      Exit Do
  '    End If
  '    i = i + 1
  '  Loop
  '  s.Fields.Append strName, adVarChar, 1200
  'End If
  s.Open
  
  
  m = UBound(gc)
  i = 0
  c = 0
  While i <= m
    strName = gc(i)
    If FieldExists(t, strName) Then
      ReDim Preserve gcTemp(c) As String
      gcTemp(c) = gc(i)
      c = c + 1
    End If
    i = i + 1
  Wend
  
  gc = gcTemp
  If IsArray(gc, True) Then
    ubGroupColumns = UBound(gc)
  Else
    ' ubGroupColumns remains == -1
  End If

  ' if SORT_BY_GC is set then sort by the group columns
  If (options And SORT_BY_GC) > 0 Then
    t.Sort = BuildRSSort(gc, t)
    Set t = CloneRecordset(t)
  End If


  Set x = New ADODB.Recordset
  x.Fields.append "SortID", adSingle
  x.Fields.append "RowID", adSingle
  x.Open
  
  t1 = Timer
  ' Add the Group Columns gc() to the summary s recordset
  c = 1
  t.MoveFirst
  Do Until t.EOF
    If ubGroupColumns < 0 Then
      strFilter = "TOTAL"
    Else
      strFilter = BuildRSFilter(gc, t.Fields)
    End If
    
    If h.Exists(strFilter) = False Then
      h.Add strFilter, h.Count + 1
    End If
    
    x.AddNew
    x.Fields("SortID") = h.Item(strFilter)
    x.Fields("RowID") = c
    
    t.MoveNext
    c = c + 1
  Loop
  x.Sort = "[SortID] , [RowID] ASC"
  'Set GroupRecordSet2 = CloneRecordset(x)
  'Exit Function
  
  t1 = Timer - t1
  t2 = Timer
  
  ' Add the Values vc() columns according the the actions ac()
  prvSortID = 0
  c = 0
  x.MoveFirst
  t.MoveFirst
  g.MoveFirst
  If trackerTable Is Nothing Then
    Set trackerTable = New HashTable ' pointless allocation... but saves having if statements each time....
  End If
  
  

  If headerText <> "" Then
  
    If ubGroupColumns < 0 Then
      strName = s.Fields(0).name
    Else
      strName = gc(0)
    End If
    s.AddNew
    s.Fields(strName) = headerText
    trackerTable.Add s.RecordCount, "Header"
    
    If (options And BLANK_AFTER) > 0 Then
      s.AddNew
      trackerTable.Add s.RecordCount, "Blank"
    End If
  End If

  Do Until x.EOF
    lngRowID = x.Fields("RowID")
    t.Move (lngRowID - t.AbsolutePosition)
      
    If x.Fields("SortID") <> prvSortID And prvSortID <> 0 Then
     ' This just removes some of the repeated action for totals
      BuildSubTotalHelper s, g, gc, c, totalText, options, trackerTable
      
      c = 0
      g.MoveNext
    End If

    ' Group Heading
    If x.Fields("SortID") <> prvSortID Then
      s.AddNew
      i = 0
      While i <= ubGroupColumns
        strName = gc(i)
        s.Fields(strName) = t.Fields(strName)
        i = i + 1
      Wend
      trackerTable.Add s.RecordCount, "Heading"
    End If
    
    s.AddNew
    For Each fld In t.Fields
      s.Fields(fld.name).Value = fld.Value
    Next
  
    
    
    prvSortID = x.Fields("SortID")
    c = c + 1
    x.MoveNext
  Loop

  ' Add a Final End of Group
  BuildSubTotalHelper s, g, gc, c, totalText, options, trackerTable
  g.MoveNext
    
  If (options And GRAND_TOTAL) > 0 Then
  
    If ubGroupColumns < 0 Then
      strName = s.Fields(0).name
    Else
      strName = gc(0)
    End If
    
    
    s.AddNew
    s.Fields(strName) = "Total"
    trackerTable.Add s.RecordCount, "GrandTotal"
    
    Set g = GroupRecordSet(r, "MyTotalColumThatShouldNeverExist9999", valueColumns, actions, options)
    For Each f In g.Fields
      s.Fields(f.name) = f.Value
    Next
    
    If (options And BLANK_AFTER) > 0 Then
      s.AddNew
      trackerTable.Add s.RecordCount, "Blank"
    End If
  End If
  
  t2 = Timer - t2
  
  'MsgBox Format(t1, "#0.0000") & " vs " & Format(t2, "#0.0000")
  StatusChange "SubTotalRecordSet " & s.RecordCount & " rows " & s.Fields.Count & " columns"
  Set SubTotalRecordSet = CloneRecordset(s)
  
End Function

Private Sub BuildSubTotalHelper(ByRef s As ADODB.Recordset, ByRef g As ADODB.Recordset, gc As Variant, ByVal rowCount As Long, Optional totalText As String = "", Optional options As SubTotalRecordSetOptionsEnum = 0, Optional ByRef trackerTable As HashTable)
  Dim i As Long
  Dim ubGroupColumns As Long
  Dim strHeadingName As String
  
  
  Dim aryTemp() As String
  Dim strName As String
  
  Dim fld As Variant
  

  If IsArray(gc, True) Then
    ubGroupColumns = UBound(gc)
    strHeadingName = gc(0)
  Else
    ubGroupColumns = -1
    strHeadingName = s.Fields(0).name
  End If
  
  If rowCount > 1 Or (options And NO_SINGLE_TOTALS) < 0 Then ' either more than one item or the flag not set.
    s.AddNew
    For Each fld In g.Fields
      If totalText = "" Then
        s.Fields(fld.name).Value = fld.Value
      ElseIf InArray(gc, fld.name) = False Then
        s.Fields(fld.name).Value = fld.Value
      End If
    Next
    If totalText <> "" Then
      s.Fields(strHeadingName) = totalText
      If (options And TOTAL_TEXT_INCLUDE_DATA) > 0 Then
        i = 0
        ReDim aryTemp(UBound(gc)) As String
        Do While i <= ubGroupColumns
          strName = gc(i)
          aryTemp(i) = g.Fields(strName)
          i = i + 1
        Loop
        s.Fields(strHeadingName) = s.Fields(strHeadingName) & " " & Join(aryTemp, ", ")
      End If
    End If
  End If
  trackerTable.Add s.RecordCount, "Total"
  
  If (options And BLANK_AFTER) > 0 Then
    s.AddNew
    trackerTable.Add s.RecordCount, "Blank"
  End If

End Sub


Function PivotAndSubTotalRecordSet(r As ADODB.Recordset, groupColumns As Variant, subTotalColumns As Variant, pivotColumns As Variant, valueColumn As String, Optional action As String, Optional headerText As String = "", Optional totalText As String = "", Optional ByVal options As SubTotalRecordSetOptionsEnum = 0, Optional ByRef trackerTable As HashTable) As ADODB.Recordset
  Dim s As ADODB.Recordset
  Dim strName As String
  Dim prvCol As String
  Dim pc() As String
  Dim cols() As String
  Dim uniqueValueColumns() As String
  
  
  
  Dim i As Long
  Dim c As Long
  Dim l As Long
  
  Set s = CloneRecordset(r)
  
  
  
  'find the names of our value columns
  c = 0
  pc = CStrArray(pivotColumns)
  s.MoveFirst
  Do Until s.EOF
    strName = CStr(s.Fields(pc(0)))
    If strName <> prvCol Then ' minor saving if sorted
      ReDim Preserve cols(c)
      cols(c) = strName
      c = c + 1
    End If
    
    prvCol = strName
    s.MoveNext
  Loop
  
  
  cols = SortStringArray(cols, True)
  i = 0
  c = 0
  l = UBound(cols)
  Do While i <= l
    strName = CStr(cols(i))
    If strName <> prvCol Then
      ReDim Preserve uniqueValueColumns(c)
      uniqueValueColumns(c) = strName
      c = c + 1
    End If
    prvCol = strName
    i = i + 1
  Loop
  
  Set s = PivotRecordSet(s, groupColumns, pivotColumns, valueColumn, action, options)
  Set s = SubTotalRecordSet(s, subTotalColumns, uniqueValueColumns, action, headerText, totalText, options, trackerTable)
  
  Set PivotAndSubTotalRecordSet = s
End Function


Private Function BuildRSFilter(strArray As Variant, flds As Fields) As String
  ' take an array of field names and a record
  ' and then build a filter statement based only on the fields in the strArray()
  
  Dim i, l As Long
  Dim strName As String
  Dim aryFilter() As String
  
  If IsArray(strArray, True) Then
    l = UBound(strArray)
    
    i = 0
    While i <= l
      strName = strArray(i)
      ReDim Preserve aryFilter(i) As String
      aryFilter(i) = "[" & strName & "] = '" & EscapeRSFilter(flds(strName).Value) & "'"
      i = i + 1
    Wend
    
    BuildRSFilter = Join(aryFilter, " AND ")
  Else
    BuildRSFilter = ""
  End If

End Function

Private Function BuildRSSort(strArray As Variant, Optional checkAgaintsRS As ADODB.Recordset = Nothing) As String
  ' take an array of field names and a record
  ' and then build a filter statement based only on the fields in the strArray()
  
  ' if recordset checkAgaintsRS is passed, then check for adVariant type columns
  ' as you cant sort against these.
  
  Dim i, l As Long
  Dim c As Long
  Dim strName As String
  Dim arySort() As String
  Dim isVariant As Boolean
  
  
  If IsArray(strArray, True) Then
    l = UBound(strArray)
    
    i = 0
    c = 0
    While i <= l
      strName = strArray(i)
      
      isVariant = False
      If Not checkAgaintsRS Is Nothing Then
        If checkAgaintsRS.Fields(strName).Type = adVariant Then
          isVariant = True
        End If
      End If
      
      If isVariant = False Then
        ReDim Preserve arySort(c) As String
        arySort(c) = "[" & strName & "]"
        c = c + 1
      End If
      i = i + 1

      
    Wend
    
    BuildRSSort = Join(arySort, ", ")
  Else
    BuildRSSort = ""
  End If

End Function

Private Function EscapeRSFilter(strValue As String) As String
  EscapeRSFilter = Replace(strValue, "'", "''")
End Function


Function AppendRecordset() As ADODB.Recordset

End Function


Function MergeRecordset(ParamArray param() As Variant) As ADODB.Recordset
  'param can be a list of recordset, or just one. (one would but pointless though
  
  Dim i As Long
  Dim l As Long
  Dim m As Long
  Dim c As Long
  
  Dim r As ADODB.Recordset
  Dim t As ADODB.Recordset
  
  Dim rsets() As Variant
  Dim fsets() As Variant
  Dim flds As Fields
  Dim fld As Field
  Dim fldName As String
  Dim typeName As String
  
  Dim f As Variant
  Dim col As Collection
  
  Dim tmpArray() As String
  Dim strArray() As String
  Dim varArray() As Variant
  
  
  i = 0
  l = 0
  While i <= UBound(param)
    If ObjectType(param(i)) = otADODB_RECORDSET Then
      ReDim Preserve rsets(l) As Variant
      ReDim Preserve fsets(l) As Variant
      Set rsets(l) = CloneRecordset(param(i))
      Set fsets(l) = rsets(l).Fields
      l = l + 1
    End If
    i = i + 1
  Wend
  
  If l < 1 Then
    Set MergeRecordset = Nothing
    Exit Function
  End If

  i = 0
  Set col = New Collection
  
  ReDim varArray(UBound(fsets)) As Variant
  While i <= UBound(fsets)
    l = 0
    ReDim tmpArray(0) As String
    Set flds = fsets(i)
    For Each f In flds
      fldName = f.name
      
      ReDim Preserve tmpArray(l) As String
      tmpArray(l) = CStr(fldName)
      
      'If InCollection(col, fldName) = True Then
      '  typeName = col.Item(fldName).Type
      '
      'End If
      
      If InCollection(col, fldName) = False Then
        col.Add f, fldName
        'MsgBox "Adding : " & fldName
      Else
        'MsgBox "NOT Adding : " & fldName
      End If
      l = l + 1
    Next
    strArray = MergeArray(strArray, tmpArray)
    i = i + 1
  Wend
  

  
  
  'MsgBox Join(strArray, ", ")
  Set r = New ADODB.Recordset

  i = 0
  While i <= UBound(strArray)
    Set fld = col(strArray(i))
    Select Case fld.Type
      Case adVariant
        r.Fields.append fld.name, fld.Type
      Case adNumeric, adDecimal
        r.Fields.append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
        r.Fields(r.Fields.Count - 1).Precision = fld.Precision
        r.Fields(r.Fields.Count - 1).NumericScale = fld.NumericScale
      Case Else
        r.Fields.append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
    End Select
    i = i + 1
  Wend
  
  r.Open


  i = 0
  While i <= UBound(rsets)
    Set t = rsets(i)
    t.MoveFirst
    Do Until t.EOF
      r.AddNew
      For Each f In t.Fields
        r.Fields(f.name).Value = t.Fields(f.name).Value ' You need to be very specific here about adding .value when using adVariant
      Next
      t.MoveNext
    Loop
    
    Set rsets(i) = Nothing ' Save a bit of memory
    i = i + 1
  Wend
  
  r.MoveFirst
  Set MergeRecordset = r
   
End Function




Function FieldsEquivalent(rs1 As ADODB.Recordset, rs2 As ADODB.Recordset) As Boolean
  ' Tests the structure of the second recordset to the first and checks if
  ' rs2 would fit into rs1. Not the othery way around.
  
  ' this needs to accept a parameter array to be really useful.
  
  
  Dim f1 As ADODB.Fields
  Dim f2 As ADODB.Fields
  Dim f As Field

  Dim fName As String
  
  Set f1 = rs1.Fields
  Set f2 = rs2.Fields
  
  If f1.Count <> f2.Count Then
    'MsgBox "different count " & f1.Count & " " & f2.Count
    FieldsEquivalent = False
    Exit Function
  End If
  
  For Each f In f1
    If FieldExists(rs2, f.name) = False Then
      'MsgBox f.name & " doesnt exist in rs2"
      FieldsEquivalent = False
      Exit Function
    End If
    
    If f.Type = adVariant Then
      ' Continue, it can take anything
    Else
      If f2(f.name).Type <> f.Type Then
        'MsgBox f.name & " have differnt types " & f2(f.name).Type & " " & f.Type
        FieldsEquivalent = False
        Exit Function
      End If
      
      If f2(f.name).DefinedSize > f.DefinedSize Then
        'MsgBox f.name & " have differnt dsize " & f2(f.name).DefinedSize & " " & f.DefinedSize
        FieldsEquivalent = False
        Exit Function
      End If


    End If
    
  Next


  ' Must be true
  FieldsEquivalent = True
  

End Function
