Sub TestaBuscaLinhaVazia()
    Dim testaLinha As Integer
    
    testaLinha = FunctionsBuscaRelatFotog.busca_ultima_linha_vazia
    
    MsgBox ("Ultima linha: " + CStr(testaLinha))
    
End Sub
