Attribute VB_Name = "SharedRecordSet"
' This module is the property of Mark Nold
' It is available for use via the LGPL license.


Option Explicit


Public Const ADT_ACTION_SUM As String = "sum"
Public Const ADT_ACTION_MAX As String = "max"
Public Const ADT_ACTION_MIN As String = "min"
Public Const ADT_ACTION_CONCATENATE As String = "concat"
Public Const ADT_ACTION_JOIN_COMMA As String = "joinc"
Public Const ADT_ACTION_JOIN_HASH As String = "joinh"
Public Const ADT_ACTION_JOIN_CRLF As String = "joinh"

Public Const VARCHAR_SIZE = 1200





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
      Recordset.Fields.Append CStr(f.name), f.Type, f.DefinedSize

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
  Dim n As ADODB.Recordset
  
  
  If r Is Nothing Then
    Set CloneEmptyRecordset = Nothing
    Exit Function
  End If

  ' This is just a cheap and nasty way of cloning a recordset then deleting everything in it
  ' i would assume s.Delete adAffectGroup would improve this but probably should build a better
  ' CloneRecordset()
    
  Set n = CloneRecordset(r.Clone)
  ' i though i could use
  ' s.Delete adAffectGroup
  ' but it appears not.
  
  'MsgBox "BEFORE " & n.RecordCount
  Do Until n.EOF Or n.RecordCount = 0
    n.MoveFirst
    'MsgBox "Delete and " & n.RecordCount & " to go"
    n.Delete
  Loop
  
  n.UpdateBatch
  'MsgBox "AFTER " & n.RecordCount
  
  Set CloneEmptyRecordset = n
End Function


Function CloneRecordsetStructure(r As ADODB.Recordset, Optional OpenOnCreate As Boolean = True) As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim oRsCloned As ADODB.Recordset
    
    Set oRsCloned = New ADODB.Recordset
    
    For Each fld In r.Fields
        oRsCloned.Fields.Append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
        
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
    Dim fld As ADODB.Field
    Dim t As ADODB.Recordset
    
    Set t = New ADODB.Recordset
    
    For Each fld In r.Fields
        t.Fields.Append fld.name, adVariant
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
    Dim fld As ADODB.Field
    Dim t As ADODB.Recordset
    
    Set t = New ADODB.Recordset
    
    For Each fld In r.Fields
        t.Fields.Append fld.name, adVarChar, VARCHAR_SIZE
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

Function CloneRecordset(ByVal oRs As ADODB.Recordset, _
  Optional ByVal LockType As ADODB.LockTypeEnum = -1) As ADODB.Recordset
  'Optional ByVal LockType As ADODB.LockTypeEnum = adLockUnspecified) As ADODB.Recordset

  ' according to http://www.vbrad.com/article.aspx?id=12
 
 
  ' Contrary to popular belief, Recordset.Clone doesn't actually clone the recordset.
  ' It doesn't actually create a new object in memory - it simply returns a reference
  ' to the original recordset with the option of making the reference read-only.
  ' To verify this claim, simply delete a record from the cloned recordset and
  ' you will see that the .RecordCount on the original recordset also decreases.
  
  ' So how do you actually make a true clone of the recordset with no dependencies
  ' or dangling references? One way is to save the recordset to file via the .Save
  ' method and then read it into another recordset. However, this method is very
  ' costly and time-consuming because ADO has to write the entire recordset structure,
  ' including field types and every property and every piece of data to disk.
  ' The proper answer is in the rarely used ADODB.Stream object.
  ' It turns out that you can save the entire recordset to this object
  ' (which is in memory) and then restore to another recordset.
  ' Check out the code below.
  
  ' Note that ADO contains a bug which shows up when using this method.
  ' When opening a recordset from a Stream object, the recordset retains
  ' all of the memory allocated to the Stream object, in addition to
  ' allocating its own. Subsequently setting the Stream object to
  ' Nothing does nothing. However, when a Recordset object is
  ' deallocated, all the memory, including that of the dead
  ' Stream object is deallocated as well.
  ' So, not a horrible bug, but something to keep in mind when
  ' you have multiple users accessing your cloning code
  ' on the web server.
  ' I've talked to the developer inside Microsoft on the
  ' ADO team who thought that this was not a bug, but rather '
  ' a design decision. Let me explain: Recordset assumes that
  ' since it was created from the IStream interface, it may be asked
  ' in the future to stream it elsewhere using this interface.
  ' Thus rather than having to recreate this interface, it simply streams
  ' the IStream memory chunk it has been created by. I think,
  ' this behavior is a bug and if it indeed was a design decision,
  ' then it was a bad design decision.
  
  ' For example, if you want to clone a really large recordset
  ' which normally takes 110 MB of RAM (11000 rows and 650 columns),
  ' you will end up "leaking" 50 MB. Your total consumption of RAM
  ' will be 270 MB (110 MB for the original RS, 110 MB for the
  ' cloned RS and 50 MB that was absorbed from the Stream object).



  ' HOWEVER (Need to test this :)
  ' Clone Method Remarks
  ' Use the Clone method to create multiple, duplicate Recordset objects, particularly if you want to be able to maintain more than one current record in a given set of records. Using the Clone method is more efficient than creating and opening a new Recordset object with the same definition as the original.
  ' The current record of a newly created clone is set to the first record.
  ' Changes you make to one Recordset object are visible in all of its clones regardless of cursor type. However, once you execute the ADO Recordset Object Requery Method on the original Recordset, the clones will no longer be synchronized to the original.
  ' Closing the original recordset does not close its copies; closing a copy does not close the original or any of the other copies.
  ' You can only clone a Recordset object that supports bookmarks. Bookmark values are interchangeable; that is, a bookmark reference from one Recordset object refers to the same record in any of its clones.




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
      r.Fields.Append fieldName, adVarChar, VARCHAR_SIZE  ' This is an Ellipse Connector hangover. Probably should type the data but stay with text for interoperability. Unfortuntaly you cant use adVarient if you want to sort and StandardText can be up to 20 lines of 60 characters.
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


Function PivotRecordSet(r As ADODB.Recordset, groupColumns As Variant, pivotColumns As Variant, valueColumn As String, Optional action As String, Optional ByVal options As GroupRecordSetOptionsEnum = 0) As ADODB.Recordset
  StatusChange "PivotRecordSet() starting"
  ' Set rt = PivotRecordSet(r, Split("ResourceGroup,ResourceID,ResourceName,ResourceType,DailyRate", ","), "ReportedWeekEnding", "DaysWorked", "sum")
  ' The order of columns seem to depend on the original order and
  ' not the specified GroupColumns
  
  
  Dim t As ADODB.Recordset
  Dim x As ADODB.Recordset
  
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
      x.Fields.Append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
      
      'special handling for data types with numeric scale & precision
      Select Case fld.Type
          Case adNumeric, adDecimal
              x.Fields(x.Fields.Count - 1).Precision = fld.Precision
              x.Fields(x.Fields.Count - 1).NumericScale = fld.NumericScale
      End Select
    End If
  Next

  
  
  
  
  valType = t.Fields(pc(0)).Type
  valSize = t.Fields(pc(0)).DefinedSize
  valAttributes = t.Fields(pc(0)).Attributes
  
  valPrecision = t.Fields(pc(0)).Precision
  valNumericScale = t.Fields(pc(0)).NumericScale
  
  
  t.MoveFirst
  Do Until t.EOF
    strName = CStr(t.Fields(pc(0)))
    If strName <> prvCol Then ' minor saving if sorted
      ReDim Preserve cols(UBound(cols) + 1)
      cols(UBound(cols)) = strName
    End If
    
    prvCol = strName
    t.MoveNext
  Loop
 
  
  
  cols = SortStringArray(cols, True)
  i = 0
  For c = 0 To UBound(cols)
    strName = CStr(cols(c))
    
    StatusChange c & " - " & strName & " " & InArray(pc, strName) & " FROM: " & Join(pc, ",")
    
    If Len(Trim(strName)) > 0 And FieldExists(x, strName) = False Then
      '       InArray(pc, strName) = False And _
       'strName <> valueColumn And _

      StatusChange "Adding field : " & strName
      x.Fields.Append strName, valType, valSize, valAttributes
      'special handling for data types with numeric scale & precision
      Select Case valType
          Case adNumeric, adDecimal
              x.Fields(x.Fields.Count - 1).Precision = valPrecision
              x.Fields(x.Fields.Count - 1).NumericScale = valNumericScale
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
    x.Filter = BuildRSFilter(gc, t.Fields)
    'x.Fields("INFO") = BuildRSFilter(gc, t.Fields)
    
    If x.RecordCount < 1 Then
      x.AddNew
      i = 0
      While i <= l
        strName = gc(i)
        x.Fields(strName) = t.Fields(strName)
        i = i + 1
      Wend
    End If
    
    i = 0
    While i <= u
      strName = t.Fields(pc(i))
      x.Fields(strName) = t.Fields(valueColumn)
      i = i + 1
    Wend
    
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
  x.Fields.Append "RowID", adSingle
  x.Fields.Append "OriginalValue", adVarChar, max
  x.Fields.Append "NumericValue", adVarChar, max + 10
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
      s.Fields.Append f.name, f.Type, f.DefinedSize, f.Attributes
      
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
      s.Fields(strName) = t.Fields(strName)
    Next
    t.MoveNext
  Loop
  
  Set DeleteFields = s
  
End Function

Function GroupRecordSet(r As ADODB.Recordset, groupColumns As Variant, valueColumns As Variant, Optional actions As Variant, Optional options As GroupRecordSetOptionsEnum = 0) As ADODB.Recordset
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
        s.Fields.Append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
    
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
        s.Fields.Append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
    
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
    t.Sort = (BuildRSSort(gc))
    Set t = CloneRecordset(t)
    'MsgBox "TRUE"
  Else
   ' MsgBox "FALSE"
  End If


  Set x = New ADODB.Recordset
  x.Fields.Append "SortID", adSingle
  x.Fields.Append "RowID", adSingle
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
        s.Fields(strName) = t.Fields(strName)
        i = i + 1
      Wend
      i = 0
      While i <= u
        strName = vc(i)
        s.Fields(strName) = t.Fields(strName)
        i = i + 1
      Wend
    Else

      i = 0
      While i <= u
        strName = vc(i)
        strAction = ac(i)
        
        Select Case strAction
        Case ADT_ACTION_SUM
          If IsNumeric(t.Fields(strName)) = True Then
            If IsNumeric(s.Fields(strName)) = True Then
              s.Fields(strName) = CDbl(s.Fields(strName)) + CDbl(t.Fields(strName))
            Else
              s.Fields(strName) = CDbl(t.Fields(strName))
            End If
            
          End If
        Case ADT_ACTION_MIN
          If IsNumeric(t.Fields(strName)) = True Then
            If CDbl(t.Fields(strName)) <= CDbl(s.Fields(strName)) Then
              s.Fields(strName) = CDbl(t.Fields(strName))
            End If
          End If
        Case ADT_ACTION_MAX
          If IsNumeric(t.Fields(strName)) = True Then
            If CDbl(t.Fields(strName)) >= CDbl(s.Fields(strName)) Then
              s.Fields(strName) = CDbl(t.Fields(strName))
            End If
          End If
        Case ADT_ACTION_CONCATENATE
          s.Fields(strName) = CStr(s.Fields(strName)) & CStr(t.Fields(strName))
        Case ADT_ACTION_JOIN_COMMA
          s.Fields(strName) = CStr(s.Fields(strName)) & "," & CStr(t.Fields(strName))
        Case ADT_ACTION_JOIN_HASH
          s.Fields(strName) = CStr(s.Fields(strName)) & "#" & CStr(t.Fields(strName))
        Case ADT_ACTION_JOIN_CRLF
          s.Fields(strName) = CStr(s.Fields(strName)) & vbCrLf & CStr(t.Fields(strName))
        
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
    t.Sort = (BuildRSSort(gc))
    Set t = CloneRecordset(t)
  End If


  Set x = New ADODB.Recordset
  x.Fields.Append "SortID", adSingle
  x.Fields.Append "RowID", adSingle
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

Private Function BuildRSSort(strArray As Variant) As String
  ' take an array of field names and a record
  ' and then build a filter statement based only on the fields in the strArray()
  
  Dim i, l As Long
  Dim strName As String
  Dim arySort() As String
  
  If IsArray(strArray, True) Then
    l = UBound(strArray)
    
    i = 0
    While i <= l
      strName = strArray(i)
      ReDim Preserve arySort(i) As String
      arySort(i) = "[" & strName & "]"
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
      ReDim Preserve tmpArray(l) As String
      tmpArray(l) = CStr(f.name)
      
      If InCollection(col, f.name) = False Then
        col.Add f, f.name
      End If
      l = l + 1
    Next
    strArray = MergeArray(strArray, tmpArray)
    i = i + 1
  Wend
  

  'strArray = MergeArrays(varArray) ' This Merges string arrays with ordering considered.
  

  Set r = New Recordset

  i = 0
  While i <= UBound(strArray)
    Set fld = col(strArray(i))
    r.Fields.Append fld.name, fld.Type, fld.DefinedSize, fld.Attributes
    Select Case fld.Type
      Case adNumeric, adDecimal
        r.Fields(r.Fields.Count - 1).Precision = fld.Precision
        r.Fields(r.Fields.Count - 1).NumericScale = fld.NumericScale
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
        'If FieldExists(r, fld.Name) = False Then
        r.Fields(f.name) = t.Fields(f.name)
        'End If
      Next
      t.MoveNext
    Loop
    i = i + 1
  Wend
  
  r.MoveFirst
  Set MergeRecordset = r
   
End Function
