Sub Insere_Agora()
    
    ' insere a data e hora atual e notificação E-mail ao enviar
    
    Dim linhaatual As Integer
    
    linhaatual = linha_Atual.linha_Atual
    
    If (Worksheets(1).Cells(linhaatual, 3) <> "") Then
    
        Worksheets(1).Cells(linhaatual, 1) = Date
        Worksheets(1).Cells(linhaatual, 2) = Time
        
        ' Worksheets(1).Cells(addr_lin, 7) = "E-mail"
        
        Functions.CreateStatusHTML
    Else
        MsgBox ("Preencha o campo numero da loja primeiro")
        
    End If
    
End Sub

