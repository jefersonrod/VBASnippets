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
    Dim name As String
    name = Mid(Application.Username, 9)
    name = Replace(name, ".", "")
    'MsgBox "Current user is " & name
    Username = name
End Function

Public Function CreateStatusHTML()
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

urlGet = urlAdd + "/add/" + col + "/" + analista + "/" + loja + "/" + dia + "/" + hora
Debug.Print urlGet

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



Debug.Print sJSONString

End Function

Public Function EmptyStatusHTML()
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

'MsgBox (col + vbCrLf + analista + vbCrLf + loja + vbCrLf + dia + vbCrLf + hora)
urlGet = urlDisp + "/disp/" + col + "/" + analista
Debug.Print urlGet

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

Debug.Print sJSONString
    
End Function

Public Function ActualSheetName() As String
Dim sheetName
sheetName = ActiveSheet.name
ActualSheetName = sheetName
End Function
Public Function CheckAnalystID() As String
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
col = Worksheets("Config.ini").Cells(linhaProcurar, colColID)
CheckAnalystID = col

End Function

Public Function CheckAnalystLine() As Integer
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
CheckAnalystLine = linhaProcurar

End Function

Public Function GetUserLogged() As String
Dim userlogged As String
userlogged = Environ$("UserName")
GetUserLogged = userlogged
End Function

Public Function checkServer(Url As String) As Boolean
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

