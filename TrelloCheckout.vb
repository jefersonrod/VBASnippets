Option Explicit

Sub TrelloCheckout(tipo As String, analista As String, lojaCodigo As String, corpoHTML As String, registrado As String)
'set email variable
Dim OutApp As Object
Dim OutMail As Object
'set general variable
Dim plan As String
Dim corpo As String
Dim email As String
Dim corpoEmail As String
Dim trello As String
Dim assunto As String
Dim emailBoardTrello As String
Dim trelloUser As String
Dim nl As Variant
Dim br As String
Dim datahoraAtual As Date
Dim gerado As String
Dim criado As String

'set config.ini vars
Dim linhaConfig As Integer
'set body vars
Dim checkin As String
Dim line As String
Dim checkout As String

line = "---"
checkout = "*Check-Out*"
'feed vars position from config.ini
emailBoardTrello = getAnalistaBoardTrello(analista)
trelloUser = getAnalistaTrelloUser(analista)
'set general var value
plan = "CheckLog"
nl = vbCrLf 'new line
br = "<br>" 'new line HTML
gerado = "Gerado em: "
criado = "Criado por: "

'Check mail address to send
    If (emailBoardTrello = "") Then
        MsgBox ("Preencha o endereço de e-mail do Board do Trello na aba CONFIG.INI do analista " + analista)
        Exit Sub
    End If
    
    'Check Trello address to add as a member
    If (trelloUser = "") Then
        MsgBox ("Preencha o endereço de usuario do Trello na aba CONFIG.INI do analista " + analista)
        Exit Sub
    End If
    

    'Feed email fields
    datahoraAtual = Now()
    assunto = "#" + lojaCodigo + " - " + tipo
    
    corpoEmail = trelloUser + br + corpoHTML + br + br + ">" + gerado + CStr(datahoraAtual) + br + ">" + criado + registrado + br + br
    'setup app definitions
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    'Build Email
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    
    'Send EMail
    With OutMail
        .To = emailBoardTrello
        '.CC = ""
        '.BCC = ""
        .Subject = assunto
        .HTMLBody = corpoEmail
        .Send
    End With

    
    'setup app definitions enable
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With


End Sub
