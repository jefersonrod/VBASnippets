Sub Cria_Email()
On Error Resume Next
On Error GoTo 0

Dim email As String
Dim trello As String
Dim assunto As String
Dim corpo As String
Dim linhaatual As Integer
Dim colLoja As Integer
Dim colData As Integer
Dim colHora As Integer
Dim colTecnico As Integer
Dim colfoneTecnico As Integer
Dim colRespLoja As Integer
Dim colfoneRespLoja As Integer
Dim colProblema As Integer
Dim colSolucao As Integer
Dim loja As String
Dim data As String
Dim hora As String
Dim tecnico As String
Dim foneTecnico As String
Dim respLoja As String
Dim foneRespLoja As String
Dim problema As String
Dim solucao As String
Dim corAzulMarinho As String


nl = vbCrLf 'new line
br = "<br>" 'new line HTML
corAzulMarinho = RGB(68, 114, 196)

'feed vars col position
colData = 1
colHora = 2
colLoja = 3
colProblema = 8
colSolucao = 9
colTecnico = 4
colfoneTecnico = 5
colRespLoja = 6
colfoneRespLoja = 7
colProblema = 8
colSolucao = 9

'get atual
linhaatual = linha_Atual.linha_Atual

'Check main fields
If (Worksheets("Atendimentos").Cells(linhaatual, colLoja) = "") Then
    MsgBox ("Favor preencher o numero da loja na coluna" + CStr(colLoja))
    Exit Sub
    
ElseIf (Worksheets("Atendimentos").Cells(linhaatual, colData) = "") Then
    MsgBox ("Favor preencher a Data na coluna" + CStr(colData))
    Exit Sub
    
ElseIf (Worksheets("Atendimentos").Cells(linhaatual, colHora) = "") Then
    MsgBox ("Favor preencher a Hora na coluna" + CStr(colHora))
    Exit Sub
    
ElseIf (Worksheets("Atendimentos").Cells(linhaatual, colProblema) = "") Then
    MsgBox ("Favor preencher o Problema na coluna" + CStr(colProblema))
    Exit Sub
    
ElseIf (Worksheets("Atendimentos").Cells(linhaatual, colSolucao) = "") Then
    MsgBox ("Favor preencher a Solução na coluna" + CStr(colSolucao))
    Exit Sub
    
Else
    loja = Worksheets("Atendimentos").Cells(linhaatual, colLoja)
    data = Worksheets("Atendimentos").Cells(linhaatual, colData)
    hora = Format(Worksheets("Atendimentos").Cells(linhaatual, colHora), "hh:mm")
    tecnico = Worksheets("Atendimentos").Cells(linhaatual, colTecnico)
    foneTecnico = Worksheets("Atendimentos").Cells(linhaatual, colfoneTecnico)
    respLoja = Worksheets("Atendimentos").Cells(linhaatual, colRespLoja)
    foneRespLoja = Worksheets("Atendimentos").Cells(linhaatual, colfoneRespLoja)
    problema = Worksheets("Atendimentos").Cells(linhaatual, colProblema)
    solucao = Worksheets("Atendimentos").Cells(linhaatual, colSolucao)
    'MsgBox (linhaAtual)
End If

'Check secondary fields
If (Worksheets("Atendimentos").Cells(linhaatual, colTecnico) <> "") Then
    tecnico = Worksheets("Atendimentos").Cells(linhaatual, colTecnico)
End If

'Check mail address to send
If (Worksheets("config.ini").Cells(1, 2) <> "") Then
    email = Worksheets("config.ini").Cells(1, 2)
Else
    MsgBox ("Preencha o endereço de e-mail na aba CONFIG.INI")
End If

'Check Trello address to add as a member
If (Worksheets("config.ini").Cells(2, 2) <> "") Then
    trello = Worksheets("config.ini").Cells(2, 2)
Else
    MsgBox ("Preencha o endereço de usuario do Trello na aba CONFIG.INI")
End If

'Feed email fields
assunto = "#" + loja + " - " + problema
corpo = trello + br + _
        "Loja: " + loja + br + _
        "Data: " + data + br + _
        "Hora: " + hora + br + _
        "Tecnico: " + tecnico + br + _
        "Fone técnico: " + foneTecnico + br + _
        "Responsavel loja: " + resLoja + br + _
        "Fone Resp. Loja: " + foneRespLoja + br + _
        "Problema: " + problema + br + _
        "Solução: " + solucao + br

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
Worksheets("Atendimentos").Cells(linhaatual, colLoja).Font.Color = corAzulMarinho
Worksheets("Atendimentos").Cells(linhaatual, colLoja).Font.Bold = True

'setup app definitions enable
With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

End Sub
