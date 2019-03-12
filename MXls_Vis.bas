Attribute VB_Name = "MXls_Vis"
Option Explicit
Function SetViszWb(A As Workbook, Vis As Boolean) As Workbook
SetViszXls A.Application, Vis
Set SetViszWb = A
End Function
Private Sub SetViszXls(A As Excel.Application, Vis As Boolean)
If A.Visible <> Vis Then A.Visible = Vis
End Sub

Function SetViszWs(A As Worksheet, Vis As Boolean) As Worksheet
SetViszXls A.Application, Vis
Set SetViszWs = A
End Function

Function WsVis(A As Worksheet) As Worksheet
XlsVis A.Application
Set WsVis = A
End Function

Function WbVis(A As Workbook) As Workbook
XlsVis A.Application
Set WbVis = A
End Function

Sub XlsVis(A As Excel.Application)
If Not A.Visible Then A.Visible = True
End Sub

Function RgVis(Rg As Range) As Range
Rg.Application.Visible = True
Set RgVis = Rg
End Function

Function LoVis(A As ListObject) As ListObject
XlsVis A.Application
Set LoVis = A
End Function

