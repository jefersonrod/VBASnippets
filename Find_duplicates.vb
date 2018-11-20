Sub Find_duplicates()
'set vars
Dim count As Integer
Dim bssBase As Integer
Dim final As Integer
Dim fimArray As Integer
Dim line As String
Dim planilhaBase As String
Dim planilhaDestino As String
Dim duplicados As String
Dim sslido As String

Dim SSs() As String
'feed that damn vars
planilhaBase = Worksheets("config.ini").Cells(2, 2)
planilhaDestino = Worksheets("config.ini").Cells(3, 2)
line = ReadAddress()
count = line
bssBase = 1
duplicados = ""
    While (Worksheets(planilhaBase).Cells(count, bssBase) <> "")
        count = count + 1
    Wend
    
final = count - line
ReDim SSs(final)
K = 0
    For I = line To count
        sslido = Worksheets(planilhaBase).Cells(I, bssBase)
        SSs(K) = sslido
        fimArray = (I - line) - 1
        For J = LBound(SSs) To fimArray
            If sslido = SSs(J) Then
                duplicados = duplicados + sslido + " - "
            End If
        Next J
        K = K + 1
    Next I
       
    If duplicados <> "" Then
        MsgBox ("Foram encontrados duplicados: " & duplicados)
    End If
    
End Sub
