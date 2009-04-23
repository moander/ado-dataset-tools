Attribute VB_Name = "SharedUtilitiesVBA"
Option Explicit


Public Sub ExportAllModules()
  Dim s As Boolean
  s = ExportVBAModules("C:\Projects\ado-dataset-tools\src.VB\")

End Sub

Function CreateDir(dirName As String) As Boolean
  Dim d() As String
  Dim i As Long
  Dim dirTemp As String
  
  CreateDir = False
  If Right(dirName, 1) <> "\" Then
    dirName = dirName & "\"
  End If

  If Dir$(dirName & "*.*", vbDirectory) = "" Then
    d = Split(dirName, "\")
    
    dirTemp = ""
    i = 0
    Do While i <= UBound(d)
      dirTemp = dirTemp & d(i) & "\"
      If Not d(i) Like "*:*" And Dir$(dirTemp & "*.*", vbDirectory) = "" Then
        MkDir (dirTemp)
      End If
      i = i + 1
      If i > 10 Then
        Exit Do
      End If
    Loop
    
  End If
  
  If Dir$(dirName & "*.*", vbDirectory) = "" Then
    'errorchange dirName & " cant be created."
    CreateDir = False
    Exit Function
  End If
  
  CreateDir = True

End Function

Function ExportVBAModules(dirName As String, Optional additionalFileExtension As String) As Boolean
  Dim appName As String
  
  Dim app As Application
  Set app = Application
  appName = app.name ' "Microsoft Excel"
  
  If appName = "Microsoft Excel" Then
    ExportVBAModules = ExportVBAExcelModules(dirName, additionalFileExtension)
  ElseIf appName = "Microsoft Project" Then
    ExportVBAModules = ExportVBAProjectModules(dirName, additionalFileExtension)
  Else
    ExportVBAModules = False
  End If
  
  
End Function



Function ExportVBAExcelModules(dirName As String, Optional additionalFileExtension As String) As Boolean
  Dim fileName As String
  Dim w As Workbook
  Dim a As AddIn
  
  
  
  Dim c As Variant
  Dim r As Variant
  Dim componentType(100) As String

  Dim s As Boolean
  
  

  
  
  
  componentType(1) = "bas"
  componentType(2) = "cls"
  componentType(3) = "frm"
  componentType(100) = "cls"
  
  
  additionalFileExtension = Trim(additionalFileExtension)
  If additionalFileExtension <> "" Then
    additionalFileExtension = "." & additionalFileExtension
  End If
    
  
  If Right(dirName, 1) <> "\" Then
    dirName = dirName & "\"
  End If
  s = CreateDir(dirName)
  
  If s = False Then
    GoTo ERRORHANDLE
  End If


  'Dim w As Variant 'Dim w As Workbook
  'Dim a As Variant 'Dim a As AddIn
  dirName = dirName & "Mircosoft Excel\"
  
  
  
  'Set p = VBAProject.ThisProject
  'Set p = ActiveWorkbook.VBProject
  For Each w In Workbooks
    For Each c In w.VBProject.VBComponents
      CreateDir (dirName & w.name)
      c.Export dirName & w.name & "\" & c.name & "." & componentType(c.Type) & additionalFileExtension
    Next
  Next
  
  
  For Each a In Excel.AddIns
    If a.name Like "*.xla" Then
      On Error Resume Next
      Set w = Excel.Workbooks(a.name)
        For Each c In w.VBProject.VBComponents
          CreateDir (dirName & w.name)
          c.Export dirName & w.name & "\" & c.name & "." & componentType(c.Type) & additionalFileExtension
        Next
      On Error GoTo 0
    End If
  Next

  
  ' It would be good to write a References.txt file to go with this
  'For Each r In p.VBProject.References
  '  MsgBox r.Name & " " & r.Type & " " & r.BuiltIn & " " & r.major & " " & r.Description
  'Next
  
  ExportVBAExcelModules = True
  Exit Function
ERRORHANDLE:
  ExportVBAExcelModules = False
  Exit Function
    
End Function

Function ExportVBAProjectModules(dirName As String, Optional additionalFileExtension As String) As Boolean
  Dim fileName As String
  Dim p As Project
  
  Dim c As Variant
  Dim r As Variant
  Dim componentType(100) As String

  Dim s As Boolean

  
  componentType(1) = "bas"
  componentType(2) = "cls"
  componentType(3) = "frm"
  componentType(100) = "cls"
  
  
  additionalFileExtension = Trim(additionalFileExtension)
  If additionalFileExtension <> "" Then
    additionalFileExtension = "." & additionalFileExtension
  End If
    
  
  If Right(dirName, 1) <> "\" Then
    dirName = dirName & "\"
  End If
  s = CreateDir(dirName)
  
  If s = False Then
    GoTo ERRORHANDLE
  End If

   
  ' "Microsoft Project"
  dirName = dirName & "Microsoft Project\"

  'Set p = VBAProject.ThisProject
  'Set p = ActiveWorkbook.VBProject
  For Each p In Projects
    For Each c In p.VBProject.VBComponents
      CreateDir (dirName & p.name)
      c.Export dirName & p.name & "\" & c.name & "." & componentType(c.Type) & additionalFileExtension
    Next
  Next
    

  
  ' It would be good to write a References.txt file to go with this
  'For Each r In p.VBProject.References
  '  MsgBox r.Name & " " & r.Type & " " & r.BuiltIn & " " & r.major & " " & r.Description
  'Next
  
  ExportVBAProjectModules = True
  Exit Function
ERRORHANDLE:
  ExportVBAProjectModules = False
  Exit Function
    
End Function

