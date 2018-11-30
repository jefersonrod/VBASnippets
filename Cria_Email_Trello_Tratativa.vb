Sub Cria_Email_Trello_Tratativa()
On Error Resume Next
On Error GoTo 0

Dim plan As String
Dim email As String
Dim trello As String
Dim assunto As String
Dim corpo As String
Dim tag As String

Dim linhaAtual As Integer
Dim colLoja As Integer
Dim colCP As Integer
Dim colCidade As Integer
Dim colUF As Integer
Dim colStatusMX As Integer
Dim colSituacao As Integer
Dim colPossibilidade As Integer
Dim colResponsavel As Integer

Dim loja As String
Dim cp As String
Dim cidade As String
Dim uf As String
Dim statusmx As String
Dim situacao As String
Dim possibilidade As String
Dim responsavel As String


Dim corAzulMarinho As String


nl = vbCrLf 'new line
br = "<br>" 'new line HTML
corAzulMarinho = RGB(68, 114, 196)
plan = "Tratativas"
tag = "#tratativas"

'feed vars col position
colLoja = 1
colCP = 2
colCidade = 3
colUF = 4
colStatusMX = 6
colSituacao = 8
colPossibilidade = 9
colResponsavel = 10

'get atual
linhaAtual = linha_Atual.linha_Atual

'Check main fields
If (Worksheets(plan).Cells(linhaAtual, colResponsavel) = "") Then
    MsgBox ("Favor preencher o responsavel na coluna" + CStr(colResponsavel))
    Exit Sub

    
Else
    loja = Worksheets(plan).Cells(linhaAtual, colLoja)
    cp = Worksheets(plan).Cells(linhaAtual, colCP)
    cidade = Worksheets(plan).Cells(linhaAtual, colCidade)
    uf = Worksheets(plan).Cells(linhaAtual, colUF)
    statusmx = Worksheets(plan).Cells(linhaAtual, colStatusMX)
    situacao = Worksheets(plan).Cells(linhaAtual, colSituacao)
    possibilidade = Worksheets(plan).Cells(linhaAtual, colPossibilidade)
    responsavel = Worksheets(plan).Cells(linhaAtual, colResponsavel)
    'MsgBox (linhaAtual)
End If


'Check mail address to send
If (Worksheets("config.ini").Cells(1, 2) <> "") Then
    email = Worksheets("config.ini").Cells(1, 2)
Else
    MsgBox ("Preencha o endereço de e-mail na aba CONFIG.INI")
End If

'Check Trello address to add as a member
If (Worksheets(plan).Cells(linhaAtual, colResponsavel) <> "") Then
    trello = responsavel
Else
    MsgBox ("Preencha o endereço de responsavel do Trello na coluna " + CStr(colResponsavel))
End If

'Feed email fields
assunto = "Trat. lj " + loja + " - " + situacao + " " + tag
corpo = responsavel + br + _
        "Responsavel: " + responsavel + br + _
        "Loja: " + loja + br + _
        "CP: " + cp + br + _
        "Cidade: " + cidade + br + _
        "UF: " + uf + br + _
        "Status MX: " + statusmx + br + _
        "Situação: " + situacao + br + _
        "Possibilidade: " + possibilidade + br


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
    .To = email
    '.CC = ""
    '.BCC = ""
    .Subject = assunto
    .HTMLBody = corpo
    .Send
End With

'change color of the loja (colunm 1) to mark as sent
Worksheets(plan).Cells(linhaAtual, colLoja).Font.Color = corAzulMarinho
Worksheets(plan).Cells(linhaAtual, colLoja).Font.Bold = True

'setup app definitions enable
With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

End Sub


