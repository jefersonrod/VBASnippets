Private Sub Workbook_BeforeClose(Cancel As Boolean)
Dim plan As String
plan = ActualSheetName 'obtem o nome da planilha atual
If ActiveSheet.name = plan Then
    Cancel = True
End If

End Sub
