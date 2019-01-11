Option Explicit
Public Function coluna_Atual() As String
'get actual line number, works only up to colunm Z
Dim addr_col As String
Dim addr_vlr As String

' Verifica a linha onde esta o cursor
    ' alimenta variaveis com coordenadas de linha e coluna
    addr_vlr = Application.ActiveCell.Address
    
    addr_col = Mid(addr_vlr, 2, 1)

    coluna_Atual = addr_col

Exit Function

End Function
