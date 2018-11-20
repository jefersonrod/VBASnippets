Sub Imprime_netsh_aspas()

    ' Chama comando netsh para o ip informado
    '
    ' Jeferson 12/12/2013
    '
    Dim IP As String
        
    
    ' informa o endereço da celula selecionada para a variavel
    ' (linha, coluna)
    addr_vlr = Application.ActiveCell.Address
    
    ' define o nome do arquivo do corpo do e-mail
    ' Arqv = "e-mail" & "_" & Cli & ".txt"
    Arqv = "teste.txt"
    
    ' define o caminho completo do arquivo de e-mail
    ' ******* Surgiu problema com função shell ao abrir o arquivo usando variavel, verificar depois
    Drive = "C:\temp\"
    FullArqv = Drive + Arqv
     
     
    
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
    IP = Worksheets(1).Cells(addr_lin, 11)
    
    ' verifica se o campo de end ip foi preenchido
    
    If IP = "" Then MsgBox ("Preencher o campo IP antes de fazer acesso remoto"): End
    
    ' manipular arquivo
    
    Open FullArqv For Output As #1
    
        'Print #1,
        'Print #1, " "
        Print #1, "netsh interface ip set address name=""Conexão local"" static " & IP & " 255.255.255.224 " & IP
   
    
    ' fecha o arquivo de texto
    Close #1
    
   
    
    
    ' abre o bloco de notas com o texto criado
    x = Shell("notepad" & " " & FullArqv, vbNormalNoFocus)

End Sub

