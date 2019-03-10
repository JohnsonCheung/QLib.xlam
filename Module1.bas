Attribute VB_Name = "Module1"
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
