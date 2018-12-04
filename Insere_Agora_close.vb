Sub Insere_Agora_close()
    
    ' insere a data e hora atual e notificação E-mail ao enviar
    
    Dim linhaAtual As Integer
    Dim plan As String
    
    linhaAtual = linha_Atual.linha_Atual 'chama função para obter a linha atual
    plan = ActualSheetName 'obtem o nome da planilha atual
    

    If (Worksheets(plan).Cells(linhaAtual, 2) = "") Then
        MsgBox ("Falta preencher a hora inicial. Verifique")
    Else
       Worksheets(plan).Cells(linhaAtual, 13) = Time
       
    End If
    
    ' Worksheets(1).Cells(addr_lin, 7) = "E-mail"
    
    EmptyStatusHTML
End Sub
