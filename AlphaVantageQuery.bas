Attribute VB_Name = "AlphaVantageQuery"
''=======================================================
'' PROGRAM: AVQ (Alpha Vantage Query) Excel-Add-In (https://github.com/LelandGrunt/AVQ)
'' VERSION: 1.0.0
'' LICENSE: MIT (https://opensource.org/licenses/MIT)
'' DESCRIPTION: Simulation of the XLQ User-Defined Function of Excel-Add-In xlquotes (Stock Prices for Microsoft Excel©).
''              Also, Excel User-Defined Function wrapper for Alpha Vantage API for financial data.
''              Supported Alpha Vantage functions: Stock Time Series Data on Daily basis (TIME_SERIES_DAILY).
'' PREREQUIREMENT: Free Alpha Vantage API Key (https://www.alphavantage.co/support/#api-key)
'' ARGUMENTS: symbol - Ticker symbol (Mandatory)
''              item - Name of item to return (Optional, Default "close")
''               day - Day/X-th item of the time series (Optional, Default 0)
''              date - Date of the time series (Optional)
'' EXAMPLES: =AVQD("MSFT")                        - Returns the recent "close" stock quote of Microsoft Corporation.
''           =AVQD("MSFT";"close")                - Returns the recent "close" stock quote of Microsoft Corporation.
''           =AVQD("MSFT";"high";-2)              - Returns the "high" stock quote of Microsoft Corporation from two day ago.
''           =AVQD("MSFT";"open";7)               - Returns the 7th "open" stock quote from the Alpha Vantage query result of Microsoft Corporation.
''           =AVQD("MSFT";"volume";;"2017-11-15") - Returns the trading volume of Microsoft Corporation of 2017-11-15.
''           =XLQ("MSFT")                         - Returns the recent "close" stock quote of Microsoft Corporation.
'' ERROR: Returns Excel error #N/A if no value was found, or #NULL! if an error occurred.
'' COMMENTS: (1) Supported item values: open, high, low, close, volume.
''           (2) If day = 0 then most recent data point is selected (Date unspecific).
''               If day < 0 then data point of current date minus day is selected (Date specific).
''               If day > 0 then the x-th (x = day) data point is selected (Date unspecific).
''           (3) If date is given then it overrules day parameter.
'' CHANGES----------------------------------------------
'' Date         Developer      Change
'' 2017-11-26   Leland Grunt   First public release.
'' -----------------------------------------------------
'' Copyright © 2017 Leland Grunt (leland.grunt[at]gmail.com)
'' Stock data are provided by Alpha Vantage (www.alphavantage.co)
'' Alpha Vantage is a copyright of Alpha Vantage Inc.
'' Alpha Vantage Terms of Service https://www.alphavantage.co/terms_of_service/
'' Uses VBA JSON Parser and Converter (https://github.com/VBA-tools/VBA-JSON)
'' xlquotes is a copyright of Dirk Voigtländer (www.xlquotes.com)
''=======================================================
Option Explicit

Public Const addInName As String = "AVQ"

'' Signatur of XLQ function.
Public Function XLQ(ticker As String, _
                    Optional side As String = "LAST", _
                    Optional hist As String = "") As Variant
    On Error GoTo Error
    
    '' Check if Alpha Vantage API Key is set, if not exit function.
    Dim apiKey As String
    If CheckApiKey = False Then
        XLQ = CVErr(xlErrNA)
        Exit Function
    Else
        apiKey = GetApiKey
    End If
    
    Static sendUnsupportedHistMessage As Boolean
    Static sendUnsupportedSideMessage As Boolean
    Dim item As String
    '' If no quote was found, then return Excel #N/A error.
    Dim quote As Variant: quote = CVErr(xlErrNA)
    
    If hist <> vbNullString Then
        '' Send notification only once in a session.
        If sendUnsupportedHistMessage = False Then
            MsgBox hist + " is not supported by AVQ.", , addInName
            sendUnsupportedHistMessage = True
        End If
        Exit Function
    ElseIf UCase(side) = "LAST" Or UCase(side) = "LETZTER" Then
        item = "close"
    ElseIf UCase(side) = "CLOSE" Or UCase(side) = "SCHLUSS" Then
        item = "close"
    ElseIf UCase(side) = "OPEN" Or UCase(side) = "ERÖFFNUNG" Then
        item = "open"
    ElseIf UCase(side) = "LOW" Or UCase(side) = "TIEF" Then
        item = "low"
    ElseIf UCase(side) = "HIGH" Or UCase(side) = "HOCH" Then
        item = "low"
    ElseIf UCase(side) = "VOLUME" Or UCase(side) = "VOLUMEN" Then
        item = "volume"
    Else
        '' Send notification only once in a session.
        If sendUnsupportedSideMessage = False Then
            MsgBox side + " is not provided by Alpha Vantage.", , addInName
            sendUnsupportedSideMessage = True
        End If
        Exit Function
    End If
    
    quote = AVQD(ticker, item)
    XLQ = quote
    
    Exit Function
    
Error:
    '' If an error occurs, then return #NULL! error.
    XLQ = CVErr(xlErrNull)
End Function

'' Alpha Vantage Query for daily equity data.
Public Function AVQD(ByVal symbol As String, _
                     Optional ByVal item As String = "close", _
                     Optional ByVal day As Integer = 0, _
                     Optional ByVal quoteDate As Date) As Variant
    On Error GoTo Error
    
    '' Check if Alpha Vantage API Key is set, if not exit function.
    Dim apiKey As String
    If CheckApiKey = False Then
        AVQD = CVErr(xlErrNA)
        Exit Function
    Else
        apiKey = GetApiKey
    End If
    
    '' Switch to wait cursor, to notificate the user about background activities.
    Application.Cursor = xlWait
    
    '' "This API returns daily time series (date, daily open, daily high, daily low, daily close, daily volume) _
        of the equity specified, covering up to 20 years of historical data."
    Rem Const apiKey As String = "<Alpha Vantage API Key>"
    Const API_FUNCTION As String = "TIME_SERIES_DAILY"
    Const URL_ALPHA_VANTAGE_QUERY As String = "https://www.alphavantage.co/query"
    
    Dim url As String
    Dim http As Object
    Dim json As Dictionary
    Dim timeSeriesDaily As Dictionary
    Dim strQuoteDate As String: strQuoteDate = ""
    Dim quoteDay As Dictionary
    '' If no quote was found, then return Excel #N/A error.
    Dim quote As Variant: quote = CVErr(xlErrNA)
    
    '' Provided Alpha Vantage Time Series Data (Default is "4. close").
    '' For ease of use, the preceding numbering is not necessary.
    Select Case item
        Case "open", "1. open"
            item = "1. open"
        Case "high", "2. high"
            item = "2. high"
        Case "low", "3. low"
            item = "3. low"
        Case "close", "4. close"
            item = "4. close"
        Case "volume", "5. volume"
            item = "5. volume"
        Case Else
            item = "4. close"
    End Select
    
    '' If optional quoteDate parameter is set, then select the data point of given quoteDate.
    If quoteDate <> "00:00:00" Then
        strQuoteDate = Format(quoteDate, "YYYY-MM-DD")
    Else
        '' Select the most recent data point if 0 (default) was given.
        If day <> 0 Then
            '' If day is negative, then select data point "current date minus <day>".
            If day < 0 Then
                strQuoteDate = Format(DateAdd("d", day, Date), "YYYY-MM-DD")
            Else '' Else select the data point at position <day> (zero-based index).
                day = day - 1
            End If
        End If
    End If
    
    '' API documentation: https://www.alphavantage.co/documentation/#daily
    url = URL_ALPHA_VANTAGE_QUERY + "?" + "function=" + API_FUNCTION + "&symbol=" + symbol + "&apikey=" + apiKey
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    With http
        .Open "GET", url, False
        .Send
        '' Get JSON text/data and convert to Dictionary object.
        Set json = JsonConverter.ParseJson(.responseText)
        If Not (json Is Nothing) Then
            '' Get JSON object "Time Series (Daily)" with time series data points.
            Set timeSeriesDaily = json("Time Series (Daily)")
            If strQuoteDate <> vbNullString Then '' Get JSON object "Date" by given <quoteDate>.
                If timeSeriesDaily.Exists(strQuoteDate) Then
                    Set quoteDay = timeSeriesDaily(strQuoteDate)
                End If
            '' Get JSON object "Date" by index "day".
            ElseIf timeSeriesDaily.Count >= day + 1 Then
                Set quoteDay = timeSeriesDaily.Items(day)
            End If
            If Not (quoteDay Is Nothing) Then
                If quoteDay.Exists(item) Then
                    quote = quoteDay.item(item)
                    '' Set quote value as Double with decimal separator which match to current Excel settings.
                    Select Case Application.International(xlDecimalSeparator)
                        Case ","
                            quote = CDbl(Replace(Replace(quote, ",", ""), ".", ","))
                        Case Else
                            quote = CDbl(quote)
                    End Select
                End If
            End If
            '' Return quote.
            AVQD = quote
            GoTo Cleanup
        End If
    End With
    
Error:
    '' If an error occurs, then return #NULL! error.
    AVQD = CVErr(xlErrNull)
    
Cleanup:
    '' Switch back to default cursor.
    Application.Cursor = xlDefault
    '' Cleanup.
    Set http = Nothing
End Function

Private Function CheckApiKey() As Boolean
    Static sendMissingApiKeyMessage As Boolean
    Dim apiKey As String
    apiKey = GetApiKey
    If apiKey = vbNullString Then
        '' Send notification only once in a session.
        If sendMissingApiKeyMessage = False Then
            MsgBox "Please set Alpha Vantage API Key." + vbNewLine + _
                   "Go to Ribbon AVQ -> Set Api Key." + vbNewLine + _
                   "Claim your free API Key here: https://www.alphavantage.co/support/#api-key", , addInName
            sendMissingApiKeyMessage = True
        End If
        CheckApiKey = False
    Else
        CheckApiKey = True
    End If
End Function

Private Function GetApiKey() As String
    Dim apiKey As String
    apiKey = Settings.Range("A1").Value2
    GetApiKey = apiKey
End Function

Private Sub SaveApiKey(ByVal apiKey As String)
    On Error Resume Next
    '' Alpha Vantage API Key is saved in "internal" worksheet "Settings".
    Settings.Range("A1").Value2 = apiKey
    If Application.Workbooks.Count > 0 Then
        Application.CalculateBeforeSave = False
    End If
    ThisWorkbook.Save
    If Application.Workbooks.Count > 0 Then
        Application.CalculateBeforeSave = True
    End If
End Sub

Public Sub SetApiKey(control As IRibbonControl)
    Dim apiKey As String
    apiKey = InputBox("Alpha Vantage API Key:", addInName)
    If apiKey <> vbNullString Then
        SaveApiKey apiKey
        Rem Application.CalculateFullRebuild
        RefreshAll Nothing
    End If
End Sub

Public Sub RefreshAll(control As IRibbonControl)
    Rem Application.CalculateFull
    If Not ActiveWorkbook Is Nothing Then
        '' Re-Calculate only AVQ UDFs.
        Cells.Replace What:="=AVQD(", Replacement:="=AVQD("
        '' Re-Calculate only XLQ UDF.
        Cells.Replace What:="=XLQ(", Replacement:="=XLQ("
    End If
End Sub

Public Sub RefreshSelection(control As IRibbonControl)
    If Not Selection Is Nothing Then
        Selection.Replace What:="=", Replacement:="="
    End If
End Sub

Private Sub OpenLink(ByVal url As String)
    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")
    '' For security reasons only opening of links are allowed.
    wshShell.Run "http://" + url

Cleanup:
    '' Cleanup.
    Set wshShell = Nothing
End Sub

Public Sub OpenHelpLink(control As IRibbonControl)
    Dim url As String
    url = control.Tag
    OpenLink (url)
End Sub

Private Sub Build()
    Dim fileName As String
    Dim fileFullName As String
    
    SaveApiKey ""
    
    ThisWorkbook.RemovePersonalInformation = False
    
    fileName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - InStrRev(ThisWorkbook.Name, ".") - 1)
    
    fileFullName = ThisWorkbook.Path + "\GitHub\" + fileName + ".xlam"
    ThisWorkbook.SaveAs fileFullName, XlFileFormat.xlOpenXMLAddIn, , , , , , , False
    
    '' Excel 2003 currently not supported.
    Rem fileFullName = ThisWorkbook.Path + "\GitHub\" + fileName + ".xla"
    Rem ThisWorkbook.SaveAs fileFullName, XlFileFormat.xlAddIn, , , , , , , False
    
    ThisWorkbook.Close
End Sub
