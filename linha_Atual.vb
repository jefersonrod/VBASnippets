Option Explicit
Public Function linha_Atual() As Integer
'get actual line number, works only up to colunm Z
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
