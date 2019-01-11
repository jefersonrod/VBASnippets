Sub ClearForm()
Dim plan As String
Dim colNum As Integer


plan = FunctionsTimeModelX.ActualSheetName
colSelect = coluna_Atual.coluna_Atual

If (colSelect = "C") Then
    'Ivan
    colNum = 3
ElseIf (colSelect = "G") Then
    'Jeferson
    colNum = 7
ElseIf (colSelect = "K") Then
    'Luiz
    colNum = 11
ElseIf (colSelect = "O") Then
    'Rener
    colNum = 15
ElseIf (colSelect = "S") Then
    'Thiago
    colNum = 19
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
