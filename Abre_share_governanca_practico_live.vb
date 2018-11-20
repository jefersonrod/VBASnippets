Sub Abre_share_governanca_practico_live()


    ' Chama share remoto para o ip informado
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
    
    If str_num_ip = "" Then MsgBox ("Preencher o campo IP antes de fazer acesso ao share"): End
    If str_num_ip = " " Then MsgBox ("Preencher o campo IP antes de fazer acesso ao share"): End
   
    ' "cmd /k " & Chr(34) &
    ' strcmd = "cmd /k " & Chr(34) & "runas /u:" & str_num_ip & "\suporte /savecred cmd /c" & Chr(34)
    
    '/ u& Chr(34) &:& Chr(34) & Chr(34) &"
    ' & str_num_ip & Chr(34) & \suporte /savecred cmd /k start \\ & Chr(34) & str_num_ip  & Chr(34) & \c\temp
    'strcmd = "This is the Test only." & Chr(34) & "double ok" & Chr(34) & "."
    
    ' abre o comando ping -t
    ' x = Shell("start" & " " & "\\" & str_num_ip & "\c\temp", vbNormalNoFocus)
    ' x = Shell("ping" & " " & str_num_ip & " " & "-t", vbNormalNoFocus)
     
     'x = Shell("ping" & " " & str_num_ip, vbNormalFocus)
     'MsgBox ("start \\" & str_num_ip)
      ' x = Shell("cmd /k runas /u:" & str_num_ip & "\suporte" & " " & "/savecred" & " cmd /k " & "start \\" & str_num_ip & "\c\temp", vbNormalFocus)
     
     'MsgBox (strcmd)
     'Z = Shell(strcmd, vbNormalFocus)
     
     'y = Shell("cmd /c time /t>c:\share-ini.txt")
     x = Shell("cmd /c start \\" & str_num_ip & "\c\SUPERDB\Governanca\PracticoLive", vbNormalFocus)
     
     'y = Shell("cmd /c time /t>c:\share-end.txt")
    
     
End Sub

