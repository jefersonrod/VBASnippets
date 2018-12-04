Option Explicit
Public Function CopyText(Text As String)
    'VBA Macro using late binding to copy text to clipboard.
    'By Justin Kay, 8/15/2014
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText Text
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Function

Public Function Username() As String
'Get username from computer
    Dim name As String
    name = Mid(Application.Username, 9) 'remove prefix
    name = Replace(name, ".", "") 'remove dot
    
    Username = name
End Function

Public Function CreateStatusHTML()
'Send values via endpoint to status server for update attendance status
'set vars
Dim nl As String
Dim linhaAtual As Integer
Dim col As String
Dim plan As String
Dim analista As String
Dim loja As String
Dim dia As String
Dim hora As String
Dim colLoja As Integer
Dim server As String
Dim port As String
Dim urlAdd As String
Dim urlGet As String
Dim sJSONString As String

'set config.ini vars
Dim linhaConfig As Integer
Dim colColID As Integer
Dim colPlan As Integer
Dim colEmailBoardTrello As Integer
Dim colTrelloUser As Integer
Dim colEmailCorp As Integer

'feed vars position from config.ini
linhaConfig = CheckAnalystLine
colColID = 1
colPlan = 2
colEmailBoardTrello = 3
colTrelloUser = 4
colEmailCorp = 5

nl = vbCrLf 'new line
linhaAtual = linha_Atual.linha_Atual
plan = ActualSheetName
colLoja = 3
server = Worksheets("config.ini").Cells(1, 4)
port = Worksheets("config.ini").Cells(1, 2)
urlAdd = "http://" + server + ":" + port
col = CheckAnalystID
analista = Username
loja = Worksheets(plan).Cells(linhaAtual, colLoja)
dia = Date
dia = Replace(dia, "/", "%2F")
hora = Time
hora = Replace(hora, ":", "%3A")

'build url get string
urlGet = urlAdd + "/add/" + col + "/" + analista + "/" + loja + "/" + dia + "/" + hora
Debug.Print urlGet

'check if url is online
If (checkServer(urlAdd)) Then
    'MsgBox ("OK")
    With CreateObject("MSXML2.XMLHTTP")
    .Open "GET", urlGet, False
    .Send
    sJSONString = .ResponseText
    End With
Else
    MsgBox ("Error 404" + nl + "Server Down" + nl + "Plz check")
End If

'Debug.Print sJSONString

End Function

Public Function EmptyStatusHTML()
'Send values via endpoint to status server for update available status
'set vars
Dim nl As String
Dim linhaAtual As Integer
Dim col As String
Dim analista As String
Dim server As String
Dim port As String
Dim urlDisp As String
Dim urlGet As String
Dim sJSONString As String

'set config.ini vars
Dim linhaConfig As Integer
Dim colColID As Integer
Dim colPlan As Integer
Dim colEmailBoardTrello As Integer
Dim colTrelloUser As Integer
Dim colEmailCorp As Integer

'feed vars position from config.ini
linhaConfig = CheckAnalystLine
colColID = 1
colPlan = 2
colEmailBoardTrello = 3
colTrelloUser = 4
colEmailCorp = 5

nl = vbCrLf 'new line
linhaAtual = linha_Atual.linha_Atual
server = Worksheets("config.ini").Cells(1, 4)
port = Worksheets("config.ini").Cells(1, 2)
urlDisp = "http://" + server + ":" + port
col = CheckAnalystID
analista = Username

'build url string
urlGet = urlDisp + "/disp/" + col + "/" + analista
Debug.Print urlGet

'check if server is available
If (checkServer(urlDisp)) Then
    'MsgBox ("OK")
    With CreateObject("MSXML2.XMLHTTP")
    .Open "GET", urlGet, False
    .Send
    sJSONString = .ResponseText
    End With
Else
    MsgBox ("Error 404" + nl + "Server Down" + nl + "Plz check")
End If

'Debug.Print sJSONString
    
End Function

Public Function ActualSheetName() As String
'Get actual sheet name
Dim sheetName
sheetName = ActiveSheet.name
ActualSheetName = sheetName    'return sheet name
End Function
Public Function CheckAnalystID() As String
'get colid from analyst
Dim analista As String
Dim linhaProcurar As Integer
Dim col As String
Dim colColID As Integer
Dim colPlan As Integer
Dim colEmailBoardTrello As Integer
Dim colTrelloUser As Integer
Dim colEmailCorp As Integer

colColID = 1
colPlan = 2
analista = ActualSheetName
linhaProcurar = 3

Do While (Worksheets("config.ini").Cells(linhaProcurar, colPlan) <> analista)
    linhaProcurar = linhaProcurar + 1
Loop
col = Worksheets("Config.ini").Cells(linhaProcurar, colColID) 'return col number
CheckAnalystID = col

End Function

Public Function CheckAnalystLine() As Integer
'get analyst config line for use colunm parameters
Dim analista As String
Dim linhaProcurar As Integer

Dim colColID As Integer
Dim colPlan As Integer
Dim colEmailBoardTrello As Integer
Dim colTrelloUser As Integer
Dim colEmailCorp As Integer

colPlan = 2
analista = ActualSheetName
linhaProcurar = 3

Do While (Worksheets("config.ini").Cells(linhaProcurar, colPlan) <> analista)
    
    linhaProcurar = linhaProcurar + 1
Loop
CheckAnalystLine = linhaProcurar 'return line value

End Function

Public Function GetUserLogged() As String
'get user logged in computer
Dim userlogged As String
userlogged = Environ$("UserName")
GetUserLogged = userlogged
End Function

Public Function checkServer(Url As String) As Boolean
'function to check url online
Dim Request As Object
Dim ff As Integer
Dim rc As Variant
    
On Error GoTo EndNow
Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
With Request
   .Open "GET", Url, False
   .Send
    rc = .StatusText
End With
Set Request = Nothing
If rc = "OK" Then checkServer = True Else checkServer = False
EndNow:
End Function

