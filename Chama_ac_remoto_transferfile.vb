Sub Chama_ac_remoto_transferfile()

    ' Chama acesso remoto para o ip informado
    '
    ' isscntr /a<address> /c<command> /l /s<core server>
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
    
    End Select
    
    ' Lê posição onde esta o numero do IP
    str_num_ip = Worksheets(1).Cells(addr_lin, 11)
    
   ' verifica se o campo de end ip foi preenchido
    
    If str_num_ip = "" Then MsgBox ("Preencher o campo IP antes de fazer acesso remoto"): End
   
   
    ' desabilita o forefront
    Z = Shell("C:\Program Files\Forefront TMG Client\fwctool.exe disable", vbHide)
    
    ' abre o bloco de notas com o texto criado
    'x = Shell("C:\Program Files\LANDesk\ManagementSuite\isscntr.exe" & " " & "/a" & str_num_ip & " " & "/crunas /u:suporte cmd", vbNormalNoFocus)
    x = Shell("C:\Program Files\LANDesk\ManagementSuite\isscntr.exe" & " " & "/a" & str_num_ip & " " & "/c""file transfer""", vbNormalNoFocus)
End Sub
