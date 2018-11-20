Sub Insere_Agora_close()
    
    ' insere a data e hora atual e notificação E-mail ao enviar
    
    Dim linhaatual As Integer
    
    linhaatual = linha_Atual.linha_Atual
    

    If (Worksheets(1).Cells(linhaatual, 2) = "") Then
        MsgBox ("Falta preencher a hora inicial. Verifique")
    Else
       Worksheets(1).Cells(linhaatual, 13) = Time
       
    End If
    
    ' Worksheets(1).Cells(addr_lin, 7) = "E-mail"
    
    Functions.EmptyStatusHTML
End Sub
