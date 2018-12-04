Sub Cria_Email()
'set error definition
On Error Resume Next
On Error GoTo 0
'set general variable
Dim plan As String
Dim email As String
Dim trello As String
Dim assunto As String
Dim corpo As String
Dim linhaAtual As Integer

'sel position var
Dim colLoja As Integer
Dim colData As Integer
Dim colHora As Integer
Dim colTecnico As Integer
Dim colfoneTecnico As Integer
Dim colRespLoja As Integer
Dim colfoneRespLoja As Integer
Dim colProblema As Integer
Dim colSolucao As Integer
'set form var
Dim loja As String
Dim Data As String
Dim hora As String
Dim Tecnico As String
Dim foneTecnico As String
Dim respLoja As String
Dim foneRespLoja As String
Dim Problema As String
Dim solucao As String
Dim corAzulMarinho As String

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

'set general var value
plan = FunctionsTimeModelX.ActualSheetName
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

'get atual line
linhaAtual = linha_Atual.linha_Atual

'check if Trello card has been created
If (Worksheets(plan).Cells(linhaAtual, colLoja).Font.Color = corAzulMarinho) Then
    MsgBox ("Este atendimento ja foi criado no Trello, verifique")
    Exit Sub
Else

    'Check main fields
    If (Worksheets(plan).Cells(linhaAtual, colLoja) = "") Then
        MsgBox ("Favor preencher o numero da loja na coluna" + CStr(colLoja))
        Exit Sub
        
    ElseIf (Worksheets(plan).Cells(linhaAtual, colData) = "") Then
        MsgBox ("Favor preencher a Data na coluna" + CStr(colData))
        Exit Sub
        
    ElseIf (Worksheets(plan).Cells(linhaAtual, colHora) = "") Then
        MsgBox ("Favor preencher a Hora na coluna" + CStr(colHora))
        Exit Sub
        
    ElseIf (Worksheets(plan).Cells(linhaAtual, colProblema) = "") Then
        MsgBox ("Favor preencher o Problema na coluna" + CStr(colProblema))
        Exit Sub
        
    ElseIf (Worksheets(plan).Cells(linhaAtual, colSolucao) = "") Then
        MsgBox ("Favor preencher a Solução na coluna" + CStr(colSolucao))
        Exit Sub
        
    Else
        'if all right, feed vars
        loja = Worksheets(plan).Cells(linhaAtual, colLoja)
        Data = Worksheets(plan).Cells(linhaAtual, colData)
        hora = Format(Worksheets(plan).Cells(linhaAtual, colHora), "hh:mm")
        Tecnico = Worksheets(plan).Cells(linhaAtual, colTecnico)
        foneTecnico = Worksheets(plan).Cells(linhaAtual, colfoneTecnico)
        respLoja = Worksheets(plan).Cells(linhaAtual, colRespLoja)
        foneRespLoja = Worksheets(plan).Cells(linhaAtual, colfoneRespLoja)
        Problema = Worksheets(plan).Cells(linhaAtual, colProblema)
        solucao = Worksheets(plan).Cells(linhaAtual, colSolucao)
        
    End If
    
    'Check secondary fields
    If (Worksheets(plan).Cells(linhaAtual, colTecnico) <> "") Then
        Tecnico = Worksheets(plan).Cells(linhaAtual, colTecnico)
    End If
    
    'Check mail address to send
    If (Worksheets("config.ini").Cells(linhaConfig, colEmailBoardTrello) <> "") Then
        email = Worksheets("config.ini").Cells(linhaConfig, colEmailBoardTrello)
    Else
        MsgBox ("Preencha o endereço de e-mail na aba CONFIG.INI")
    End If
    
    'Check Trello address to add as a member
    If (Worksheets("config.ini").Cells(linhaConfig, colTrelloUser) <> "") Then
        trello = Worksheets("config.ini").Cells(linhaConfig, colTrelloUser)
    Else
        MsgBox ("Preencha o endereço de usuario do Trello na aba CONFIG.INI")
    End If
    
    'Feed email fields
    assunto = "#" + loja + " - " + Problema
    corpo = trello + br + _
            "Loja: " + loja + br + _
            "Data: " + Data + br + _
            "Hora: " + hora + br + _
            "Técnico: " + Tecnico + br + _
            "Fone técnico: " + foneTecnico + br + _
            "Responsável loja: " + resLoja + br + _
            "Fone Resp. Loja: " + foneRespLoja + br + _
            "Problema: " + Problema + br + _
            "Solução: " + solucao + br + br
    
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
End If

End Sub
