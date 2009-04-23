VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FROM http://www.bytemycode.com/snippets/snippet/251/
' Lenny Grover - http://www.bytemycode.com/members/member/lgrover/
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal bytes As Long)
 

' default values
Const DEFAULT_HASHSIZE = 1024
Const DEFAULT_LISTSIZE = 2048
Const DEFAULT_CHUNKSIZE = 1024
 

Private Type SlotType
    key As String
    Value As Variant
    nextItem As Long      ' 0 if last item
End Type
 

' for each hash code this array holds the first element
' in slotTable() with the corresponding hash code
Dim hashTbl() As Long
' the array that holds the data
Dim slotTable() As SlotType

' pointer to first free slot
Dim FreeNdx As Long

' size of hash table
Dim m_HashSize As Long
' size of slot table
Dim m_ListSize As Long
' chunk size
Dim m_ChunkSize As Long
' items in the slot table
Dim m_Count As Long

' This keeps the keys in the order they were entered for calls to the Keys property
Dim m_Keys As Collection
 

' member variable for IgnoreCase property
Private m_IgnoreCase As Boolean

' True if keys are searched in case-unsensitive mode
' this can be assigned to only when the hash table is empty

Property Get IgnoreCase() As Boolean
    IgnoreCase = m_IgnoreCase
End Property

Property Let IgnoreCase(ByVal newValue As Boolean)
    If m_Count Then
        Err.Raise 1001, , "The Hash Table isn't empty"
    End If
    m_IgnoreCase = newValue
End Property
 

Private Sub ExpandSlotTable(ByVal numEls As Long)
    Dim newFreeNdx As Long, i As Long
    newFreeNdx = UBound(slotTable) + 1
   
    ReDim Preserve slotTable(0 To UBound(slotTable) + numEls) As SlotType
    ' create the linked list of free items
    For i = newFreeNdx To UBound(slotTable)
        slotTable(i).nextItem = i + 1
    Next
    ' overwrite the last (wrong) value
    slotTable(UBound(slotTable)).nextItem = FreeNdx
    ' we now know where to pick the first free item
    FreeNdx = newFreeNdx
End Sub
 

Private Function HashCode(key As String) As Long
    Dim lastEl As Long, i As Long
   
    ' copy ansi codes into an array of long
    lastEl = (Len(key) - 1) \ 4
    ReDim codes(lastEl) As Long
    ' this also converts from Unicode to ANSI
    CopyMemory codes(0), ByVal key, Len(key)
   
    ' XOR the ANSI codes of all characters
    For i = 0 To lastEl
        HashCode = HashCode Xor codes(i)
    Next
   
End Function
 

' get the index where an item is stored or 0 if not found
' if Create = True the item is created
'
' on exit Create=True only if a slot has been actually created

Private Function GetSlotIndex(ByVal key As String, Optional Create As Boolean, Optional HCode As Long, Optional LastNdx As Long) As Long
    Dim ndx As Long
   
    ' raise error if invalid key
    If Len(key) = 0 Then Err.Raise 1001, , "Invalid key"
   
    ' keep case-unsensitiveness into account
    If m_IgnoreCase Then key = UCase$(key)
    ' get the index in the hashTbl() array
    HCode = HashCode(key) Mod m_HashSize
    ' get the pointer to the slotTable() array
    ndx = hashTbl(HCode)
   
    ' exit if there is no item with that hash code
    Do While ndx
        ' compare key with actual value
        If slotTable(ndx).key = key Then Exit Do
        ' remember last pointer
        LastNdx = ndx
        ' check the next item
        ndx = slotTable(ndx).nextItem
    Loop
   
    ' create a new item if not there
    If ndx = 0 And Create Then
        ndx = GetFreeSlot()
        PrepareSlot ndx, key, HCode, LastNdx
    Else
        ' signal that no item has been created
        Create = False
    End If
    ' this is the return value
    GetSlotIndex = ndx

End Function
 

' return the first free slot

Private Function GetFreeSlot() As Long
    ' allocate new memory if necessary
    If FreeNdx = 0 Then ExpandSlotTable m_ChunkSize
    ' use the first slot
    GetFreeSlot = FreeNdx
    ' update the pointer to the first slot
    FreeNdx = slotTable(GetFreeSlot).nextItem
    ' signal this as the end of the linked list
    slotTable(GetFreeSlot).nextItem = 0
    ' we have one more item
    m_Count = m_Count + 1
End Function

' assign a key and value to a given slot

Private Sub PrepareSlot(ByVal Index As Long, ByVal key As String, ByVal HCode As Long, ByVal LastNdx As Long)
    ' assign the key
    ' keep case-sensitiveness into account
    If m_IgnoreCase Then key = UCase$(key)
    slotTable(Index).key = key
   
    If LastNdx Then
        ' this is the successor of another slot
        slotTable(LastNdx).nextItem = Index
    Else
        ' this is the first slot for a given hash code
        hashTbl(HCode) = Index
    End If
End Sub
 

Private Sub Class_Initialize()
    ' initialize the tables at default size
    SetSize DEFAULT_HASHSIZE, DEFAULT_LISTSIZE, DEFAULT_CHUNKSIZE
    Set m_Keys = New Collection
End Sub
 

' initialize the hash table

Sub SetSize(ByVal HashSize As Long, Optional ByVal ListSize As Long, Optional ByVal ChunkSize As Long)
    ' provide defaults
    If ListSize <= 0 Then ListSize = m_ListSize
    If ChunkSize <= 0 Then ChunkSize = m_ChunkSize
    ' save size values
    m_HashSize = HashSize
    m_ListSize = ListSize
    m_ChunkSize = ChunkSize
    m_Count = 0
    ' rebuild tables
    FreeNdx = 0
    ReDim hashTbl(0 To HashSize - 1) As Long
    ReDim slotTable(0) As SlotType
    ExpandSlotTable m_ListSize
End Sub
 

' check whether an item is in the hash table

Function Exists(key As String) As Boolean
    Exists = GetSlotIndex(key) <> 0
End Function

' add a new element to the hash table

Sub Add(key As String, Value As Variant)
    Dim ndx As Long, Create As Boolean
   
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    Create = True
    ndx = GetSlotIndex(key, Create)
   
    If Create Then
        ' the item was actually added
        If IsObject(Value) Then
            Set slotTable(ndx).Value = Value
        Else
            slotTable(ndx).Value = Value
        End If
       
        m_Keys.Add key
    Else
        ' raise error "This key is already associated with an item of this collection"
        Err.Raise 457
    End If
End Sub
 

' Sort the values of the HashTable
Sub TESTListValues()
   
    'yQuickSort 1, m_Keys.Count
    'MsgBox m_Keys.item(1)
    
    Dim i As Integer
    Dim myValues() As Variant
    Dim myKeys() As Variant
    
    ReDim myKeys(m_Keys.Count - 1)
    ReDim myValues(m_Keys.Count - 1)
    
    For i = 0 To m_Keys.Count - 1
      myValues(i) = slotTable(i + 1).Value
      myKeys(i) = m_Keys.Item(i + 1)
    Next i
    
    MsgBox m_Keys.Count & " recs, " & Join(myKeys, ", ") & " with values: " & Join(myValues, ", ")
    
End Sub

' the value associated to a key
' (empty if not found)

Property Get Item(key As String) As Variant

    Dim ndx As Long
    ' get the index to the slot where the value is
    ndx = GetSlotIndex(key)
    If ndx = 0 Then
        ' return Empty if not found
    ElseIf IsObject(slotTable(ndx).Value) Then
        Set Item = slotTable(ndx).Value
    Else
        Item = slotTable(ndx).Value
    End If
End Property

Property Let Item(key As String, Value As Variant)
    Dim ndx As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    ndx = GetSlotIndex(key, True)
    ' store the value
    slotTable(ndx).Value = Value
End Property

Property Set Item(key As String, Value As Object)
    Dim ndx As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    ndx = GetSlotIndex(key, True)
    ' store the value
    Set slotTable(ndx).Value = Value
End Property
 

' remove an item from the hash table

Sub Remove(key As String)
    Dim ndx As Long, HCode As Long, LastNdx As Long
    Dim i As Integer
   
    ndx = GetSlotIndex(key, False, HCode, LastNdx)
    ' raise error if no such element
    If ndx = 0 Then Err.Raise 5
   
    If LastNdx Then
        ' this isn't the first item in the slotTable() array
        slotTable(LastNdx).nextItem = slotTable(ndx).nextItem
    ElseIf slotTable(ndx).nextItem Then
        ' this is the first item in the slotTable() array
        ' and is followed by one or more items
        hashTbl(HCode) = slotTable(ndx).nextItem
    Else
        ' this is the only item in the slotTable() array
        ' for this hash code
        hashTbl(HCode) = 0
    End If
   
    ' put the element back in the free list
    slotTable(ndx).nextItem = FreeNdx
    FreeNdx = ndx
   
    ' Remove the item from the keys collection
    For i = m_Keys.Count To 1 Step -1
        If m_Keys.Item(i) = key Then
            m_Keys.Remove (i)
        End If
    Next i
   
    ' we have deleted an item
    m_Count = m_Count - 1
   
End Sub

' remove all items from the hash table

Sub RemoveAll()
    SetSize m_HashSize, m_ListSize, m_ChunkSize
   
    ' Clear the keys collection
    Set m_Keys = New Collection
End Sub
 

' the number of items in the hash table

Property Get Count() As Long
    Count = m_Count
End Property

' the array of all keys
' (VB5 users: convert return type to Variant)

Property Get Keys() As Variant
    Dim res() As Variant
    Dim i As Integer
   
    ReDim res(m_Keys.Count - 1)
    For i = 0 To m_Keys.Count - 1
        res(i) = m_Keys.Item(i + 1)
    Next i
   
    Keys = res()
End Property
