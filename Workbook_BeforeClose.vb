Private Sub Workbook_BeforeClose(Cancel As Boolean)

If ActiveSheet.Name = "Atendimentos Migração V-Sat.xlsm" Then

    Cancel = True
    
End If

End Sub
