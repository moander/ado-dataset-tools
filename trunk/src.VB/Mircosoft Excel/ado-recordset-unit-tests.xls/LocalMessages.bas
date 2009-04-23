Attribute VB_Name = "LocalMessages"
Option Explicit

Public Function StatusChange(msg As String)
  Dim shortMsg As String
  Dim longMsg As String
  msg = Trim(msg)
  
  If Len(msg) > 50 Then
    shortMsg = Left(msg, 50) & "..."
  Else
    shortMsg = msg
  End If
  
  If Len(msg) > 200 Then
    longMsg = Left(msg, 200) & "..."
  Else
    longMsg = msg
  End If
  
  
  If msg = "" Then
    Application.DisplayStatusBar = True
    Application.StatusBar = ""
  Else
    Application.DisplayStatusBar = True
    Application.StatusBar = "Status : " & shortMsg
  End If
End Function


Public Function ErrorChange(msg As String)
  StatusChange "ERROR: " & msg
  MsgBox "ERROR: " & msg
 
End Function


Public Sub DebugChange(msgType As String, msg As String)
   
End Sub
