Attribute VB_Name = "QXls_TreeWs_Install"
Option Compare Text
Option Explicit
Private Const CMod$ = "MXls_TreeWs_Install."
Private Const Asm$ = "QXls"
Function TreeWsMdLines$()
Erase XX
X "Option Explicit"
X "Private Sub Worksheet_Change(ByVal Target As Range)"
X "MTreeWs.Change Target"
X "End Sub"

X "Private Sub Worksheet_SelectionChange(ByVal Target As Range)"
X "MTreeWs.SelectionChange Target"
X "End Sub"
TreeWsMdLines = JnCrLf(XX)
Erase XX
End Function

Sub InstallTreeWs()
Dim Ws, WB
For Each Ws In Itr(TreeWsAy)
    InstallTreeWsz CvWs(Ws)
Next
For Each WB In Itr(TreeWbAy)
    InstallTreeWbz CvWb(WB)
Next
End Sub
Sub InstallTreeWbz(WB As Workbook)
Exit Sub
'Wb.VBProject.References.AddFromFile ShpCstFxa
End Sub
Function IsTreeWb(A As Workbook) As Boolean
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If Ws.Name = "TreeWs" Then IsTreeWb = True: Exit Function
Next
End Function
Function TreeWbAy() As Workbook()
Dim WB As Workbook, Ws As Worksheet
For Each WB In Xls.Workbooks
    If IsTreeWb(WB) Then PushObj TreeWbAy, WB
Next

End Function
Function TreeWsAy() As Worksheet()
Dim WB As Workbook, Ws As Worksheet
For Each WB In Xls.Workbooks
    For Each Ws In WB.Sheets
        If Ws.Name = "TreeWs" Then PushObj TreeWsAy, Ws
    Next
Next
End Function

Sub InstallTreeWsz(A As Worksheet)
Dim Md As CodeModule
Set Md = MdzWs(A)
Stop
If Md.CountOfLines = 0 Then
    Md.AddFromString TreeWsMdLines
    InfLin CSub, "TreeWs in Wb is installed with code", "Wb", WbnzWs(A)
Else
    InfLin CSub, "TreeWs in Wb already has code", "Wb", WbnzWs(A)
End If
End Sub

