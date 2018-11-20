Option Explicit
Public Function ReadAddress() As String

  Dim addr_vlr As String
  Dim addr_line As String

  'informa o endereço da celula selecionada para a variavel
  '(linha coluna)
  'feed the address of cell selected to variable
  ' (line, Column)
addr_vlr = Application.ActiveCell.Address

'  Verifica a linha onde esta o cursor
' Verify line where is cursor

Select Case Len(addr_vlr)

Case 4
' 1 decimal position
addr_line = Int(Right(addr_vlr, 1))
Case 5
' 2 decimal position
addr_line = Int(Right(addr_vlr, 2))
Case 6
' 3 decimal position
addr_line = Int(Right(addr_vlr, 3))
Case 7
' 4 decimal position
addr_line = Int(Right(addr_vlr, 4))

End Select

ReadAddress = addr_line
Exit Function

End Function
