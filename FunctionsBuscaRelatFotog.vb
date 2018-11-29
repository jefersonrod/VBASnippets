Option Explicit
Public Function linha_Atual() As Integer
'usar : variavelTipoInt = linha_Atual.linha_Atual
Dim addr_lin As Integer
Dim addr_vlr As String

' Verifica a linha onde esta o cursor
    ' alimenta variaveis com coordenadas de linha e coluna
    addr_vlr = Application.ActiveCell.Address
    
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

    linha_Atual = addr_lin

Exit Function

End Function

Public Function busca_ultima_linha_vazia_log() As Integer
Dim linhaInicial As Integer
Dim config As String
Dim colunaBase As Integer

linhaInicial = 2
colunaBase = 1
config = "Log"
While (Worksheets(config).Cells(linhaInicial, colunaBase) <> "")
        linhaInicial = linhaInicial + 1
Wend
busca_ultima_linha_vazia_log = linhaInicial
End Function

Public Function qtde_linhas() As Integer
Dim count As Integer
Dim stLine As Integer
Dim ltLine As Integer

stLine = FunctionsBuscaRelatFotog.linha_Atual
ltLine = FunctionsBuscaRelatFotog.busca_ultima_linha_vazia_buscar 'search for last line empty
qtde_linhas = ltLine - stLine

End Function

Public Function busca_ultima_linha_vazia_buscar() As Integer
Dim linhaInicial As Integer
Dim plan As String
Dim colunaBase As Integer

linhaInicial = FunctionsBuscaRelatFotog.linha_Atual
colunaBase = 1
plan = "Buscar"
While (Worksheets(plan).Cells(linhaInicial, colunaBase) <> "")
        linhaInicial = linhaInicial + 1
Wend
busca_ultima_linha_vazia_buscar = linhaInicial
End Function
