Attribute VB_Name = "SharedErrors"
Option Explicit
' Use sparingly, since this is a library logic should handle
' all eventualities, and in the case of Recordsets pass back Nothing


Enum UserErrorEnum
  ' reserve 400 - 500 for http
  [_First] = 1
  eHTTPerror = 1
  eDataNotFound = 2
  
  eRecordsetEmpty = 3
  
  eUnknown = 4
  [_Last] = eUnknown
End Enum


Function UErrDesc(userErrorNumber As UserErrorEnum, Optional extraDetails As String = "") As String

  Dim d(UserErrorEnum.[_Last]) As String
  Dim n As Long
  
  d(eDataNotFound) = "Data not found at specified location"
  d(eHTTPerror) = "Unspecified HTTP Error"
  
  d(eRecordsetEmpty) = "Recordset Empty."
  d(eUnknown) = "Explicitly Unknown cause of error."
  

  For n = UserErrorEnum.[_First] To UserErrorEnum.[_Last]
    If d(n) = "" Then
      d(n) = "Undescribed User Defined Error"
    End If
    d(n) = d(n) & vbCrLf & _
    extraDetails & vbCrLf & _
    " Err.Number = " & UErrNumber(n)
  Next n

  UErrDesc = d(userErrorNumber)

End Function

Function UErrNumber(userErrorNumber As UserErrorEnum) As Long
  ' Error Numbers (from err.raise):
  ' Long integer that identifies the nature of the error.
  ' Visual Basic errors (both Visual Basic-defined and user-defined errors)
  ' are in the range 0–65535. The range 0–512 is reserved for system errors;
  ' the range 513–65535 is available for user-defined errors.
  ' When setting the Number property to your own error code in a class module,
  ' you add your error code number to the vbObjectError constant.
  ' For example, to generate the error number 513, assign vbObjectError + 513 to the Number property.
  UErrNumber = vbObjectError + 512
End Function

