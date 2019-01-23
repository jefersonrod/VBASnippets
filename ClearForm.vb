Sub ClearForm()
Dim plan As String
Dim colNum As Integer
Dim usuarioAtual As String
Dim ivan As String
Dim jeferson As String
Dim luiz As String
Dim rener As String
Dim thiago As String


ivan = "Ivan Erison Gambarra da Silva"
jeferson = "Jeferson Rodrigues"
luiz = "Luiz Victor Lomba de Oliveira"
rener = "Renervaldo Wizenffat"
thiago = "Thiago Hiroshi Da Silva Endo"
plan = FunctionsTimeModelX.ActualSheetName
colSelect = coluna_Atual.coluna_Atual
registrado = FunctionsTimeModelX.Username
usuarioAtual = FunctionsTimeModelX.Username


If (colSelect = "C" And usuarioAtual = ivan) Then
    analista = "Ivan"
    colNum = 3
ElseIf (colSelect = "G" And usuarioAtual = jeferson) Then
    analista = "Jeferson"
    colNum = 7
ElseIf (colSelect = "K" And usuarioAtual = luiz) Then
    analista = "Luiz"
    colNum = 11
ElseIf (colSelect = "O" And usuarioAtual = rener) Then
    analista = "Rener"
    colNum = 15
ElseIf (colSelect = "S" And usuarioAtual = thiago) Then
    analista = "Thiago"
    colNum = 19
Else
    MsgBox ("Ola " + usuarioAtual + " coluna selecionada " + colSelect + " é invalida, verifique")
    Exit Sub
End If

Dim resultado As VbMsgBoxResult
     resultado = MsgBox("Tem certeza que deseja limpar o formulário?", vbYesNo, "Limpar formulário " + plan)
     If resultado = vbYes Then
        If (plan = "Checkin") Then
        
            Worksheets(plan).Cells(11, colNum) = ""
            Worksheets(plan).Cells(13, colNum) = ""
            Worksheets(plan).Cells(15, colNum) = ""
            Worksheets(plan).Cells(17, colNum) = ""
            Worksheets(plan).Cells(19, colNum) = ""
            Worksheets(plan).Cells(23, colNum) = ""
            Worksheets(plan).Cells(25, colNum) = ""
            Worksheets(plan).Cells(27, colNum) = ""
            Worksheets(plan).Cells(29, colNum) = ""
            Worksheets(plan).Cells(31, colNum) = ""
            Worksheets(plan).Cells(5, colNum) = ""
            Worksheets(plan).Cells(9, colNum) = ""
            
        ElseIf (plan = "Checkout") Then
            
            Worksheets(plan).Cells(13, colNum) = ""
            Worksheets(plan).Cells(15, colNum) = ""
            Worksheets(plan).Cells(17, colNum) = ""
            Worksheets(plan).Cells(5, colNum) = ""
            Worksheets(plan).Cells(9, colNum) = ""
        End If
     Else
        'nothing to do
     End If

End Sub
