Attribute VB_Name = "SharedUtilitiesGoogle"
Option Explicit


' There is a google API for finance http://code.google.com/apis/finance/
' however it states : "Note that this API interacts only with users' portfolio data; the Finance Portfolio Data API doesn't provide access to stock quotes or other real-world financial data hosted by Google Finance."
' So it's a bit useless here.

' a possible list of NASDAQ NYSE and AMEX could be found http://www.nasdaq.com/asp/symbols.asp?exchange=Q&start=G&Type=0

' See http://www.eoddata.com/Symbols/AMEX/B.htm

Const DEFAULT_EXCHANGE As String = "NASDAQ"

Function GetGoogleFinanceData(companySymbol As String, startDate As Date, endDate As Date, Optional weekly As Boolean = False, Optional datesAsUniversal As Boolean = False, Optional companyDesc As String = "") As ADODB.Recordset
  
  Dim companyName As String
  Dim txt As String
  Dim host As String
  Dim url As String
  Dim qid As String
  Dim histPeriod As String
  
  
  Dim r As ADODB.Recordset
  If Trim(companySymbol) = "" Then
    Set GetGoogleFinanceData = Nothing
    Exit Function
  End If
  
  Dim sDate As String
  Dim eDate As String
  host = "www.google.com"
  
  ' qid = "NASDAQ:GOOG" or something similar
  If InStr(companySymbol, ":") < 1 Then
    ' NB: Goggle doesnt server hostory via CSV for non NASDAQ, NYSE, AMEX companies like "ASX:OZL"
    qid = DEFAULT_EXCHANGE & ":" & companySymbol
  Else
    qid = companySymbol
  End If
  
  If Trim(companyDesc) = "" Then
    companyName = qid
  Else
    companyName = Trim(companyDesc)
  End If
  
  If weekly = True Then
    histPeriod = "histperiod=weekly&" ' null implies daily. There is only daily and weekly
  End If
  
  sDate = Format(startDate, "mmm dd, yyyy")
  eDate = Format(endDate, "mmm dd, yyyy")
  
  sDate = URLEncode(sDate, True)
  eDate = URLEncode(eDate, True)
        
  url = "/finance/historical?" & _
        histPeriod & _
        "q=" & qid & _
        "&startdate=" & sDate & _
        "&enddate=" & eDate & _
        "&output=csv"
        
        
  txt = GetURL(host, url)
  
  ' expect fields: Date,Open,High,Low,Close,Volume with a \n line delimeter. ie: Chr(&HA)
  ' This would be more elegant if it handled 404 etc...
  ' but it is only example code.
  If Left(txt, 31) <> "Date,Open,High,Low,Close,Volume" Then
    Err.Raise UErrNumber(eDataNotFound), "GetGoogleFinanceData", UErrDesc(eDataNotFound, url)
    Set GetGoogleFinanceData = Nothing
    Exit Function
  End If
  
  
  
  Set r = CSVToRecordSet(txt, datesAsUniversal)
  Set r = AddFields(r, "CompanyName", companyName, False)
  Set GetGoogleFinanceData = r
  
End Function
