Sub Insere_Agora()
    
    ' insere a data e hora atual e notificação E-mail ao enviar
    
    Dim linhaAtual As Integer
    Dim plan As String
    
    linhaAtual = linha_Atual.linha_Atual 'chama função para obter a linha atual
    plan = ActualSheetName 'obtem o nome da planilha atual
    
    If (Worksheets(plan).Cells(linhaAtual, 3) <> "") Then
    
        Worksheets(plan).Cells(linhaAtual, 1) = Date
        Worksheets(plan).Cells(linhaAtual, 2) = Time
        
        ' Worksheets(1).Cells(addr_lin, 7) = "E-mail"
        
        CreateStatusHTML
    Else
        MsgBox ("Preencha o campo numero da loja primeiro")
        
    End If
    
End Sub

