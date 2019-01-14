Option Explicit
Public Function ultimaLinhaLivreChkLog() As Integer
Dim linhaInicial As Integer
Dim colunaBase As Integer
Dim plan As String

linhaInicial = 2
colunaBase = 1
plan = "CheckLog"

While (Worksheets(plan).Cells(linhaInicial, colunaBase) <> "")
        linhaInicial = linhaInicial + 1
Wend
ultimaLinhaLivreChkLog = linhaInicial
End Function
