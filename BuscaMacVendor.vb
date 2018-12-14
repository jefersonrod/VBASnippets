Sub BuscaMacVendor()
Dim macAddress As String
Dim url As String
Dim urlMacAddress As String
Dim linhaAtual As Integer
Dim plan As String

urlMacAddress = "https://api.macvendors.com/"
plan = FunctionsTimeModelX.ActualSheetName
linhaAtual = linha_Atual.linha_Atual
macAddress = Worksheets(plan).Cells(linhaAtual, 15)
macAddress = Replace(macAddress, ":", "-")
url = urlMacAddress + macAddress

If (macAddress = "" Or macAddress = " ") Then
    MsgBox ("numero do macAddress esta vazio verifique!")
Else
    'MsgBox (GetAPIMAC(url))
    Worksheets(plan).Cells(linhaAtual, 16) = GetAPIMAC(url)
End If
End Sub
