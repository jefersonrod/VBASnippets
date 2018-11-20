Sub Loja_info()
    
    Dim addr_vlr As String
    
     ' informa o endereço da celula selecionada para a variavel
    ' (linha, coluna)
    addr_vlr = Application.ActiveCell.Address
    
    ' Verifica a linha onde esta o cursor
    
    Select Case Len(addr_vlr)
    
    Case 4
    addr_lin = Int(Right(addr_vlr, 1))
    Case 5
    addr_lin = Int(Right(addr_vlr, 2))
    Case 6
    addr_lin = Int(Right(addr_vlr, 3))
    Case 7
    addr_lin = Int(Right(addr_vlr, 4))
    
    End Select
    
    
    atend_loja = Worksheets(1).Cells(addr_lin, 3)
    
    
    
    ' Abre checklist para gerar atendimento
    y = Shell("C:\Program Files\Google\Chrome\Application\chrome.exe http://sab.bull.com.br/scripts/cgiip.exe/WService=wsbroker1/LOJA021.P?k=2629&url=&pas=", vbNormalFocus)
    
    ' tempo de pausa
    Application.Wait (Now + TimeValue("0:00:03"))
    
    ' envia tab ate chegar no campo da loja
    ' SendKeys "{TAB}"
      
    
    ' envia numero da loja
    SendKeys atend_loja
    SendKeys "{tab}"
    SendKeys "{enter}"
    
    
    
End Sub
