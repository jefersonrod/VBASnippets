Sub Ping_eterno()

    ' Efetua ping continuo no ip informado
    '
    '
    Dim str_num_ip As String
    
    
    ' Seta variavel com posição atual do cursor
    
    addr_vlr = Application.ActiveCell.Address
    
    ' Verifica a linha onde esta o cursor
    ' alimenta variaveis com coordenadas de linha e coluna
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
    
    ' Lê posição onde esta o numero do IP
    str_num_ip = Worksheets(1).Cells(addr_lin, 11)
    
     ' verifica se o campo de end ip foi preenchido
    
    If str_num_ip = "" Then MsgBox ("Preencher o campo IP antes de fazer acesso remoto"): End
   
    
    ' abre o comando ping -t
    x = Shell("ping" & " " & str_num_ip & " " & "-t", vbNormalNoFocus)
    
End Sub