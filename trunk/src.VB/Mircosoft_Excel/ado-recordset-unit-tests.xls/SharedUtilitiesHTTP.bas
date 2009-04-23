Attribute VB_Name = "SharedUtilitiesHTTP"
' This module is the property of Enspace Pty Ltd (markn@enspace.com)
'
' Unfortunatly Excel Add In dont allow for easy operation between Add Ins
' so all modules have been bundled together.
' Users are expected to respect the copyright holders of these modules
' and unnauthorised distribution is expressly forbidden.
 

Option Explicit

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" ( _
    ByVal lpszAgent As String, _
    ByVal dwAccessType As Long, _
    ByVal lpszProxyName As String, _
    ByVal lpszProxyBypass As String, _
    ByVal dwFlags As Long) As Long

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" ( _
    ByVal hInternetSession As Long, _
    ByVal lpszServerName As String, _
    ByVal nServerPort As Integer, _
    ByVal lpszUsername As String, _
    ByVal lpszPassword As String, _
    ByVal dwService As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long
    
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" ( _
    ByVal hHttpSession As Long, _
    ByVal lpszVerb As String, _
    ByVal lpszObjectName As String, _
    ByVal lpszVersion As String, _
    ByVal lpszReferer As String, _
    ByVal lpszAcceptTypes As String, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long

Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" ( _
    ByVal hHttpRequest As Long, _
    ByVal lpszHeaders As String, _
    ByVal dwHeadersLength As Long, _
    ByVal lpOptional As String, _
    ByVal dwOptionalLength As Long) As Boolean


Private Declare Function InternetReadFile Lib "wininet.dll" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal dwNumberOfBytesToRead As Long, _
    ByRef lpNumberOfBytesRead As Long) As Boolean

Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_SERVICE_HTTP = 3

Const INTERNET_FLAG_NO_COOKIES = &H80000
Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000



Public Function GetURL(strServer As String, _
                        strURL As String, _
                        Optional iPort As Integer = 80, _
                        Optional strUsername As String = "", _
                        Optional strPassword As String = "") As String
  ' If you were to declare this public it could be dangerous
  ' would suggest some wrapper functions to make it usuable in Excel generally
  
  ' 1. Need interface for other variables
  ' 2. Need Error checking... look at status, file not found etc
  ' 3. Method for turning \n to \r\n
  
  Dim hInternet As Long
  Dim hConnect As Long
  Dim lFlags As Long
  Dim hRequest As Long
  Dim bRes As Boolean
  Dim strBuffer As String * 1 ' A 1 byte buffer.. slow but safe
  Dim strResult As String
  Dim lBytesRead As Long
  
  'MsgBox "connecting to: " & strServer & strURL & ":" & iPort
   
  hInternet = InternetOpen(Application.name, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
  hConnect = InternetConnect(hInternet, strServer, iPort, "", "", INTERNET_SERVICE_HTTP, 0, 0)

  lFlags = INTERNET_FLAG_NO_COOKIES
  lFlags = lFlags Or INTERNET_FLAG_NO_CACHE_WRITE
  hRequest = HttpOpenRequest(hConnect, "GET", strURL, "HTTP/1.0", vbNullString, vbNullString, lFlags, 0)

  bRes = HttpSendRequest(hRequest, vbNullString, 0, vbNullString, 0)

  Do
      bRes = InternetReadFile(hRequest, strBuffer, Len(strBuffer), lBytesRead)
      If lBytesRead > 0 Then
          strResult = strResult & strBuffer
      End If
  Loop While lBytesRead > 0
  GetURL = strResult
End Function


Sub OpenBrowserURL(strURL As String)
    Application.FollowHyperlink _
      Address:=strURL, NewWindow:=True
    
End Sub


Public Function URLEncode(StringToEncode As String, Optional UsePlusRatherThanHexForSpace As Boolean = False) As String
  ' This mimics "iso-8859-1" encoding
  ' This is problematic, becuase your server may use something different.
  ' you could force it by reading RawUrl in c#
  ' Encoding enc = Encoding.GetEncoding("iso-8859-1");
  ' System.Web.HttpUtility.UrlDecode(myString,enc)
  
  Dim TempAns As String
  Dim CurChr As Integer
  CurChr = 1
  Do Until CurChr - 1 = Len(StringToEncode)
    Select Case Asc(Mid(StringToEncode, CurChr, 1))
      'Case 37, 43, 48 To 57, 65 To 90, 97 To 122
      Case 48 To 57, 65 To 90, 97 To 122
        TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
      Case 32
        If UsePlusRatherThanHexForSpace = True Then
          TempAns = TempAns & "+"
        Else
          TempAns = TempAns & "%" & Hex(32)
        End If
     Case Else
         TempAns = TempAns & "%" & _
                Lpad( _
                LCase(Hex(Asc(Mid(StringToEncode, _
                CurChr, 1)))) _
                , "0", 2) ' %B should be %0b, %02 should remain %02. This is hard to do with Format()
  End Select
  
    CurChr = CurChr + 1
  Loop
  
  URLEncode = TempAns
End Function


Public Function URLDecode(StringToDecode As String) As String

  Dim TempAns As String
  Dim CurChr As Integer
  Dim lenStr As Long
  
  CurChr = 1
  lenStr = Len(StringToDecode)
  
  Do Until CurChr - 1 >= lenStr
    Select Case Mid(StringToDecode, CurChr, 1)
      Case "+"
        TempAns = TempAns & " "
      Case "%"
        TempAns = TempAns & Chr(val("&h" & _
           Mid(StringToDecode, CurChr + 1, 2)))
         CurChr = CurChr + 2
        '%25
      Case Else
        TempAns = TempAns & Mid(StringToDecode, CurChr, 1)
    End Select
  
  CurChr = CurChr + 1
  Loop
  
  URLDecode = TempAns
End Function
