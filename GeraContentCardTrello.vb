Sub GeraContentCardTrello()
On Error Resume Next
On Error GoTo 0

Dim plan As String
Dim corpo As String
Dim linhaAtual As Integer
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
Dim Data As String
Dim hora As String
Dim Tecnico As String
Dim foneTecnico As String
Dim respLoja As String
Dim foneRespLoja As String
Dim Problema As String
Dim solucao As String
Dim corAzul As String
Dim corAzulMarinho As String
Dim corAzulMarinhoCode As String


plan = FunctionsTimeModelX.ActualSheetName
nl = vbCrLf 'new line
corAzul = RGB(105, 134, 206) 'blue color done action
corAzulMarinho = RGB(68, 114, 196) 'marine color for card created


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
linhaAtual = linha_Atual.linha_Atual

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
        loja = Worksheets(plan).Cells(linhaAtual, colLoja)
        Data = Worksheets(plan).Cells(linhaAtual, colData)
        hora = Format(Worksheets(plan).Cells(linhaAtual, colHora), "hh:mm")
        Tecnico = Worksheets(plan).Cells(linhaAtual, colTecnico)
        foneTecnico = Worksheets(plan).Cells(linhaAtual, colfoneTecnico)
        respLoja = Worksheets(plan).Cells(linhaAtual, colRespLoja)
        foneRespLoja = Worksheets(plan).Cells(linhaAtual, colfoneRespLoja)
        Problema = Worksheets(plan).Cells(linhaAtual, colProblema)
        solucao = Worksheets(plan).Cells(linhaAtual, colSolucao)
        'MsgBox (linhaAtual)
    End If
    
    'Check secondary fields
    If (Worksheets(plan).Cells(linhaAtual, colTecnico) <> "") Then
        Tecnico = Worksheets(plan).Cells(linhaAtual, colTecnico)
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
    
    'generate content text
    
    corpo = "Loja: " + loja + nl + _
            "Data: " + Data + nl + _
            "Hora: " + hora + nl + _
            "Técnico: " + Tecnico + nl + _
            "Fone técnico: " + foneTecnico + nl + _
            "Responsável loja: " + respLoja + nl + _
            "Fone Resp. Loja: " + foneRespLoja + nl + _
            "Problema: " + Problema + nl + _
            "Solução: " + solucao + nl
    
    'generate content
    CopyText corpo 'call copy to clipboard function
    
    'change color of the loja (colunm 1) to mark as sent
    Worksheets(plan).Cells(linhaAtual, colLoja).Font.Color = corAzul




End Sub


