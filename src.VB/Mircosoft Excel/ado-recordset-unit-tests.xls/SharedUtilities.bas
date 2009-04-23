Attribute VB_Name = "SharedUtilities"
Option Explicit

' Requires reference to "Microsoft ActiveX Data Object Library 2.7"

Enum ObjectTypeEnum
     otNOTHING = 0
     otADODB_RECORDSET = 1
     otFORMATTEDRECORDSET = 2
     otFRSCELLFORMAT = 3
     otFRSFONT = 4
     otFRSBORDER = 5
     otFRSPATTERN = 6
     
     otCOLLECTION = 50
     otUNKNOWN = 99
End Enum


Function Rpad(myString As String, padString As String, padLength As Long) As String
  Dim l As Long
  l = Len(myString)
  If l > padLength Then
    padLength = l
  End If

  Rpad = Left$(myString & String(padLength, padString), padLength)
End Function

Function Lpad(myString As String, padString As String, padLength As Long) As String
  Dim l As Long
  l = Len(myString)
  If l > padLength Then
    padLength = l
  End If

  Lpad = Right$(String(padLength, padString) & myString, padLength)
End Function


Function Ceiling(Number As Double) As Long
    Ceiling = -Int(-Number)
End Function

Function SortArray(varArray() As Variant) As Variant()
  Call QuickSort(varArray, LBound(varArray), UBound(varArray))
  
  SortArray = varArray
End Function

Function SortStringArray(strArray() As String, Optional sortAsNumeric As Boolean = False) As String()

  If sortAsNumeric = True Then
    Call QuickSortS(strArray, LBound(strArray), UBound(strArray))
  Else
    Call QuickSort(strArray, LBound(strArray), UBound(strArray))
  End If
  
  SortStringArray = strArray
End Function

Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub

Function CDblOrStr(var As Variant) As Variant

  If IsNumeric(var) Then
    CDblOrStr = CDbl(var)
  Else
    CDblOrStr = var
  End If

End Function

Public Sub QuickSortS(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim compare   As Variant
  
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi
  
  'If IsNumeric(vArray((inLow + inHi) \ 2)) Then
    pivot = CDblOrStr(vArray((inLow + inHi) \ 2))
  'Else
  '  pivot = vArray((inLow + inHi) \ 2)
  'End If

  While (tmpLow <= tmpHi)
     While (CDblOrStr(vArray(tmpLow)) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < CDblOrStr(vArray(tmpHi)) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSortS vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSortS vArray, tmpLow, inHi

End Sub
Function CStrArray(str As Variant) As String()
  ' Converts a varient to a String Array if it is a string or an array
  
  Dim strArray() As String
  Dim i As Long
  Dim l As Long
 
  If IsArray(str) Then
    i = 0
    l = UBound(str)
    
    ReDim Preserve strArray(l)
    While i <= l
      strArray(i) = CStr(str(i))
      i = i + 1
    Wend
    CStrArray = strArray
    
  Else
    ReDim strArray(0)
    strArray(0) = CStr(str)
    CStrArray = strArray
  End If

End Function

Function InArray(myArray As Variant, myValue As Variant) As Boolean
  Dim i, l As Long
  
  InArray = False
  
  If IsArray(myArray, True) Then
    l = UBound(myArray)
    i = 0
    While i <= l
      If myArray(i) = myValue Then
        InArray = True
      End If
      i = i + 1
    Wend

  End If

End Function

Function IsArray(myArray As Variant, Optional andNotEmpty As Boolean) As Boolean
  IsArray = False
  Dim m As Long
  m = -1
  
  
  If andNotEmpty = False Then
    ' This tests if it's an array or not.
    ' unitialised arrays will return true.
    If VarType(myArray) >= vbArray Then
      IsArray = True
    End If
  Else
    On Error Resume Next
      m = UBound(myArray)
    On Error GoTo 0
    If m < 0 Then
      IsArray = False
    Else
      IsArray = True
    End If
  End If

End Function

Public Function InCollection(col As Collection, key As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long
  
  InCollection = False
  Set var = Nothing
  
  Err.Clear
  On Error Resume Next
    var = col.Item(key)
    errNumber = CLng(Err.Number)
  On Error GoTo 0
  
 
  
  If errNumber = 5 Then ' it's 5 if not in collection
    InCollection = False
  ElseIf errNumber = 0 Or errNumber = 438 Then ' for some reason 438...
    InCollection = True
  Else
    'MsgBox errNumber & " - " & Err.Number & " " & Err.Description
  End If
  
  

  
End Function


Private Function MergeArray_ArrayToCollection(strArray() As String) As Collection
  ' Array to Collection helper for MergeArray
  
  Dim col As Collection
  Dim it As ArrayIndex
  Dim i As Long
  Dim m As Long
  
  
  Set col = New Collection

  If VarType(strArray) <> vbArray + vbString Then
    Set MergeArray_ArrayToCollection = col
    Exit Function
  End If


  m = -1
  On Error Resume Next
  m = UBound(strArray)
  On Error GoTo 0
  If m < 0 Then
    Set MergeArray_ArrayToCollection = col
    Exit Function
  End If
  
  
  m = UBound(strArray)
  While i <= m
    Set it = New ArrayIndex
    it.Index = i
    it.Value = strArray(i)
    
    If i > 0 Then
      it.afterValue = strArray(i - 1)
    End If
      
    If i < m Then
      it.beforeValue = strArray(i + 1)
    End If
    
    it.startDistance = i
    it.endDistance = m - i
    
    If InCollection(col, it.Value) = False Then
      col.Add it, it.Value
    End If
    i = i + 1
  Wend
  
  Set MergeArray_ArrayToCollection = col

End Function

Private Function MergeArray_CollectionToArray(col As Collection) As String()
  ' Collection  to Array to  helper for MergeArray
  Dim res() As String

  Dim it As ArrayIndex
  Dim c As Variant
  Dim i As Long
  

  i = 0
  For Each c In col
    ReDim Preserve res(i) As String
    Set it = c
    res(i) = it.Value
    i = i + 1
  Next
  MergeArray_CollectionToArray = res
  
End Function

Public Function MergeArray(a1() As String, a2() As String) As String()
  Const INDEX_STEP As Double = 0.001

  Dim res() As String
  Dim sortList() As String
 
  Dim col1 As Collection
  Dim col2 As Collection
  
  Dim afValue As String
  Dim bfValue As String
  
  Dim it As ArrayIndex
  Dim c As Variant
  
  Dim i As Long
  Dim m As Long

  
  Set col1 = MergeArray_ArrayToCollection(a1)
  Set col2 = MergeArray_ArrayToCollection(a2)

  If col1.Count < 1 Then
    MergeArray = a2
    Exit Function
  End If
  
  
  For Each c In col2
    Set it = c
    
    If InCollection(col1, CStr(it.Value)) = True Then
      col2.Remove (it.Value)
    End If
  Next
  
  If col2.Count < 1 Then
    MergeArray = a1
    Exit Function
  End If
  

 
  ' If a new column exists before an existing column put it before
  ' If not then check if exists after an existing column, and put it after
  ' if not put it at the end
  
  ' NB: This doesnt support infantesimal merging due to the 0.001 (INDEX_STEP)
  ' you could take the two values and divide, but we can avoid that for now.
  '
  i = col2.Count
  For i = col2.Count To 1 Step -1
    Set it = col2(i)
    If InCollection(col1, it.beforeValue) = True Then
      'MsgBox "adding " & it.value
      it.Index = col1.Item(it.beforeValue).Index - INDEX_STEP
      it.afterValue = ""
    Else
      it.Index = col1.Count ' index is 0 based.
      it.beforeValue = ""
      it.afterValue = col1.Item(col1.Count).Value
      'col2.Remove (it.value)
      'col1.Add it, it.value
    End If
    col2.Remove (it.Value)
    col1.Add it, it.Value
    

  Next
  
  
'  If col2.Count > 1 Then
'    For Each c In col2
'      Set it = c
'      If InCollection(col1, it.afterValue) = True Then
'        it.index = col1.item(it.afterValue).index + INDEX_STEP
'        it.beforeValue = ""
'
'        col2.Remove (it.value)
'        col1.Add it, it.value
'
'      Else
'        it.index = col1.Count ' index is 0 based.
'        it.beforeValue = ""
'        it.afterValue = col1.item(col1.Count).value
'
'        col2.Remove (it.value)
'        col1.Add it, it.value
'
'      End If
'    Next
'  End If
  
  
  ' TOO LAZY TO WRITE A DECENT SORT
  ReDim sortList(col1.Count - 1) As String
  i = 0
  For Each c In col1
    Set it = c
    sortList(i) = it.Index
    i = i + 1
  Next
  Set col2 = New Collection
  sortList = SortStringArray(sortList, True)
  
  i = 0
  While i <= UBound(sortList)
    For Each c In col1
      Set it = c
      If CStr(sortList(i)) = CStr(it.Index) Then
        col2.Add it, it.Value
        col1.Remove (it.Value)
      End If
    Next
    i = i + 1
  Wend
  ' END OF TOO LAZY TO WRITE A DECENT SORT
  
  res = MergeArray_CollectionToArray(col2)

  MergeArray = res
  
End Function

Public Function MergeArrays(ParamArray param() As Variant) As String()
  Dim i As Long
  Dim m As Long
  Dim l As Long
  Dim c As Long
  
  Dim myArrays() As Variant
  Dim varArray() As Variant
  Dim myResult() As String
  Dim strArray() As String
  
  i = 0
  l = 0
  m = UBound(param)
  While i <= m
    If VarType(param(i)) = vbArray + vbString Then
      ReDim Preserve myArrays(l) As Variant
      
      myArrays(l) = param(i)
      l = l + 1
    ElseIf VarType(param(i)) >= vbArray Then
      ReDim Preserve myArrays(l) As Variant
      varArray = param(i)
      ReDim strArray(UBound(varArray)) As String
      c = 0
      While c <= UBound(strArray)
        'MsgBox VarType(varArray)
        strArray(c) = CStr(varArray(c))
        c = c + 1
      Wend
      
      myArrays(l) = strArray
      l = l + 1
    End If
    i = i + 1
  Wend
  
  If l > 0 Then
    'ReDim myResult(UBound(myArrays(0))) As String
    myResult = myArrays(0)
    i = 1
    While i <= m
      strArray = myArrays(i)
      myResult = MergeArray(myResult, strArray)
      i = i + 1
    Wend
  ElseIf l = 0 Then
    myResult = myArrays(0)
  End If
  
  MergeArrays = myResult

End Function


















Function ObjectType(obj As Variant) As ObjectTypeEnum
  Dim oType As ObjectTypeEnum
  oType = otNOTHING

  If IsObject(obj) = False Then
    Exit Function
  Else
    oType = otUNKNOWN
  End If
  
  On Error Resume Next
  oType = TestObjectType_RS(obj)
  oType = TestObjectType_FRS(obj)
  oType = TestObjectType_Cllct(obj)
  oType = TestObjectType_Frscf(obj)
  oType = TestObjectType_Frsf(obj)
  oType = TestObjectType_Frscfb(obj)
  oType = TestObjectType_Frsp(obj)
  
  On Error GoTo 0
  
  ObjectType = oType
End Function

Private Function TestObjectType_RS(ByVal obj As ADODB.Recordset) As ObjectTypeEnum
  TestObjectType_RS = otADODB_RECORDSET
End Function


Private Function TestObjectType_Cllct(ByVal obj As Collection) As ObjectTypeEnum
  TestObjectType_Cllct = otCOLLECTION
End Function


Private Function TestObjectType_FRS(ByVal obj As FormattedRecordSet) As ObjectTypeEnum
  TestObjectType_FRS = otFORMATTEDRECORDSET
End Function

Private Function TestObjectType_Frscf(ByVal obj As FRSCellFormat) As ObjectTypeEnum
  TestObjectType_Frscf = otFRSCELLFORMAT
End Function

Private Function TestObjectType_Frsf(ByVal obj As FRSFont) As ObjectTypeEnum
  TestObjectType_Frsf = otFRSFONT
End Function

Private Function TestObjectType_Frscfb(ByVal obj As FRSBorder) As ObjectTypeEnum
  TestObjectType_Frscfb = otFRSBORDER
End Function


Private Function TestObjectType_Frsp(ByVal obj As FRSPattern) As ObjectTypeEnum
  TestObjectType_Frsp = otFRSPATTERN
End Function

