'Functions_old - page 1
Option Explicit
Sub SingleHC ()
addr_line = Read Address()
if addr_line = 1 Then MsgBox ("Selected Cell is first row, don't use this") : End
Dim sUrl As String
sUrl = Worksheets ("HC").Cells(addr_line, 3)
if sUrl = "" Then MsgBox ("Please fill the URL field to check") : End
Dim ORequest As WinHttp. WinHttpRequest
Dim sResult As String
on Error GoTo Err_DoSomeJob
set oRequest = New WinHttp. WinHttpRequest
With oRequest
    .Open "GET", surl, True
    .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8" 
    .Send ""
    .WaitForResponse
    sResult = .ResponseText
    Debug.Print (sResult)
    'MsgBox (sResult)
    sResult = oRequest.Status 
    Debug.Print (sResult)
    'MsgBox (sResult) 
    Worksheets("HC").Cells(addr_line, 7) = sResult
    InsertDateHour 
    If Worksheets("HC").Cells (addr_line, 7) = 401  Then Worksheets ("HC").Cells(addr_line, 6) = "OK"
    If Worksheets ("HC").Cells (addr_line, 7) = 200 Then Worksheets ("HC"). Cells (addr line, 6) = "OK"
    If Worksheets ("HC"). Cells (addr line, 7) = 404 Then Worksheets ("HC"). Cells (addr line, 6) = "NOK"
  
End With
EXit_DoSomeJob:
  On Error Resume Next 
  Set oRequest = Nothing 
  Exit Sub
Err_DoSomeJob:
    'MsgBox Err.Description, vbExclamation, Err.Number 
    ErrorString = Err.Description
    If Right$(ErrorString, 2) = vbCrLf Or Right$(ErrorString, 2) = vbNewLine Then
              ErrorString = Left$(ErrorString, Len(ErrorString) - 2) 
            End If 
    Worksheets ("HC").Cells(addr line, 8) = ErrorString 
    InsertDateHour
    Resume Exit_DoSomeJob
End Sub
Sub Multi HC()
Application.ScreenUpdating = False
  cell_line = Read Address()
  Do
       If Worksheets("HC").Cells(cell line, 1) = "X" Or Worksheets ("HC").Cells(cell_line, 1) = "x"
  Then
      'Do Nothing 
      Else
        Call SingleHC
      End If
'Functions_old - page 2
    Active Cell.Offset (1).Select
    cell_line = cell_line + 1 
    Loop Until Worksheets("HC").Cells(cell line, 3) = "" 
    'Application.Screen Updating = True 
End Sub
'Routines - Page 1
Sub ClearCells()
cell_clear = Read Address ()
  Worksheets("HC").Cells(cell_clear, 8) = " " 
  Worksheets("HC").Cells(cell_clear, 9) = " "
  ActiveCell.Offset(1).Select 
  cell_clear = cell_clear + 1 
  Loop Until Worksheets("HC").Cells(cell_clear, 2) = ""
End Sub 
Function ReadAddress() As String
  Dim addr_vlr As String 
  Dim addr_lin As String
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
addr_line = Int(Right (addr_vlr, 1)) 
Case 5
' 2 decimal position 
addr_line = Int(Right (addr_vlr, 2)) 
Case 6
' 3 decimal position 
addr_line = Int(Right (addr_vlr, 3)) 
Case 7
' 4 decimal position 
addr_line = Int(Right (addr_vlr, 4)) 
End Select
ReadAddress = addr_line
Exit Function
End Function 
Sub InsertDateHour()
addr_line = ReadAddress()
'insere a data e hora atual 
'insert actual date and hour
Worksheets("HC").Cells(addr_line, 8) = Date 
Worksheets("HC").Cells(addr_line, 9) = Time
End Sub 
Sub GoURL()
  Dim run_url As String 
  Dim apptoload As String 
  addr_line = ReadAddress() 
  run_url = Worksheets("HC").Cells(addr_line, 5) 
  apptoload = "C:\Program Files\Internet Explorer \iexplore " 
  go = Shell(apptoload & run_url, vbNormalFocus)
End Sub
Sub Button6_Click ( )
  MsgBox "To be implemented", VbOKOnly, "Generate Report"
End Sub
Sub About()
' Routines - 2
Const strDate As String = "June/2015." 
Const strTitle. As String = "Health Check Tool."
Const strText As String = "Version 1.0"
Const strDev As String = "Jeferson R."
vx = MsgBox (strText + wbCrLif + strDate + wbCrLf + strDev, whoExclamation,
End Sub
