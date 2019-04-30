Attribute VB_Name = "MXls_Lo_Minx"
Option Explicit
Sub MinxLo(A As ListObject)
If FstTwoChr(A.Name) <> "T_" Then Exit Sub
Dim R1 As Range
Set R1 = A.DataBodyRange
If R1.Rows.Count >= 2 Then
    RgRR(R1, 2, R1.Rows.Count).EntireRow.Delete
End If
End Sub

Private Sub MinxLozWs(A As Worksheet)
If A.CodeName = "WsIdx" Then Exit Sub
If FstTwoChr(A.CodeName) <> "Ws" Then Exit Sub
Dim L As ListObject
For Each L In A.ListObjects
    MinxLo L
Next
End Sub

Function MinxLozWszWb(A As Workbook) As Workbook
Dim Ws As Worksheet
For Each Ws In A.Sheets
    MinxLozWs Ws
Next
Set MinxLozWszWb = A
End Function

Sub MinxLozWszFx(Fx$)
ClsWbNoSav SavWb(MinxLozWszWb(WbzFx(Fx)))
End Sub
