Sub FormCheckout()
'get info from sheet and generate card in Trello
On Error Resume Next
On Error GoTo 0

'set general vars
Dim plan As String
Dim corpo As String
Dim colNum As Integer
Dim colSelect As String
Dim Loja As String
Dim lojaCodigo As String
Dim line As String
Dim checkin As String
Dim checkout As String
'Técnico
Dim tecnico As String
Dim aPrimesysDeuasOrientacoes As String
Dim aPrimesysDeuasOrientacoesRESP As String
'Responsavel
Dim resposavel As String
Dim casoHajaumProblemaOcorrido As String
Dim casoHajaumProblemaOcorridoRESP As String
Dim solicitarUmaAvaliacaodoTecnico As String
Dim solicitarUmaAvaliacaodoTecnicoRESP As String
Dim seHouverMaisItens As String
Dim seHouverMaisItensRESP As String

'set general vars
plan = FunctionsTimeModelX.ActualSheetName
colSelect = coluna_Atual.coluna_Atual
nl = vbCrLf 'new line
line = "---------------------------------"
checkin = "*Check-In*"
checkout = "*Check-Out*"


If (colSelect = "C") Then
    'Ivan
    colNum = 3
ElseIf (colSelect = "G") Then
    'Jeferson
    colNum = 7
ElseIf (colSelect = "K") Then
    'Luiz
    colNum = 11
ElseIf (colSelect = "O") Then
    'Rener
    colNum = 15
ElseIf (colSelect = "S") Then
    'Thiago
    colNum = 19
    
Else
    MsgBox ("Coluna selecionada " + colSelect + " é invalida, verifique")
    Exit Sub
End If

    'Check loja field
    If (Worksheets(plan).Cells(5, colNum) = "") Then
        MsgBox ("Favor preencher o numero da loja na linha 5")
        Worksheets(plan).Cells(5, colNum).Select
        Exit Sub
        
    Else
        'check tecnico answer
        If (Worksheets(plan).Cells(9, colNum) = "") Then
            MsgBox ("Favor preencher a linha 9")
            Worksheets(plan).Cells(9, colNum).Select
            Exit Sub
        End If
        
        
        'check responsavel field
        For i = 13 To 17 Step 2
            If (Worksheets(plan).Cells(i, colNum) = "") Then
            MsgBox ("Favor preencher a linha " + CStr(i))
            Worksheets(plan).Cells(i, colNum).Select
            Exit Sub
            End If
        Next i
        
        Loja = Worksheets(plan).Cells(4, colNum)
        lojaCodigo = Worksheets(plan).Cells(5, colNum)
        'Técnico
        tecnico = Worksheets(plan).Cells(7, colNum)
        aPrimesysDeuasOrientacoes = Worksheets(plan).Cells(8, colNum)
        aPrimesysDeuasOrientacoesRESP = Worksheets(plan).Cells(9, colNum)
        'Responsavel
        responsavel = Worksheets(plan).Cells(11, colNum)
        casoHajaumProblemaOcorrido = Worksheets(plan).Cells(12, colNum)
        casoHajaumProblemaOcorridoRESP = Worksheets(plan).Cells(13, colNum)
        solicitarUmaAvaliacaodoTecnico = Worksheets(plan).Cells(14, colNum)
        solicitarUmaAvaliacaodoTecnicoRESP = Worksheets(plan).Cells(15, colNum)
        seHouverMaisItens = Worksheets(plan).Cells(16, colNum)
        seHouverMaisItensRESP = Worksheets(plan).Cells(17, colNum)
        
    End If
    
    'generate content text
    
    corpo = checkout + nl + _
            Loja + lojaCodigo + nl + nl + _
            tecnico + nl + nl + _
            aPrimesysDeuasOrientacoes + nl + _
            aPrimesysDeuasOrientacoesRESP + nl + nl + _
            line + nl + _
            responsavel + nl + nl + _
            casoHajaumProblemaOcorrido + nl + _
            casoHajaumProblemaOcorridoRESP + nl + _
            solicitarUmaAvaliacaodoTecnico + nl + _
            solicitarUmaAvaliacaodoTecnicoRESP + nl + _
            seHouverMaisItens + nl + _
            seHouverMaisItensRESP + nl
            
    
    'generate content
    CopyText corpo 'call copy to clipboard function
    
    'change color of the loja (colunm 1) to mark as sent
    'Worksheets(plan).Cells(linhaAtual, colLoja).Font.Color = corAzul


End Sub
