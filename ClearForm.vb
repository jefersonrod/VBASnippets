Sub ClearForm()
Dim plan As String
plan = FunctionsTimeModelX.ActualSheetName
Dim resultado As VbMsgBoxResult
     resultado = MsgBox("Tem certeza que deseja limpar o formulário?", vbYesNo, "Limpar formulário " + plan)
     If resultado = vbYes Then
        If (plan = "Checkin") Then
        
            Worksheets(plan).Cells(11, 3) = ""
            Worksheets(plan).Cells(13, 3) = ""
            Worksheets(plan).Cells(15, 3) = ""
            Worksheets(plan).Cells(17, 3) = ""
            Worksheets(plan).Cells(19, 3) = ""
            Worksheets(plan).Cells(23, 3) = ""
            Worksheets(plan).Cells(25, 3) = ""
            Worksheets(plan).Cells(27, 3) = ""
            Worksheets(plan).Cells(29, 3) = ""
            Worksheets(plan).Cells(5, 3) = ""
            Worksheets(plan).Cells(9, 3) = ""
            
        ElseIf (plan = "Checkout") Then
            
            Worksheets(plan).Cells(13, 3) = ""
            Worksheets(plan).Cells(15, 3) = ""
            Worksheets(plan).Cells(17, 3) = ""
            Worksheets(plan).Cells(5, 3) = ""
            Worksheets(plan).Cells(9, 3) = ""
        End If
     Else
        'nothing to do
     End If

End Sub
