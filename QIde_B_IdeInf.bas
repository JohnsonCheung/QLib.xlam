Attribute VB_Name = "QIde_B_IdeInf"
Option Explicit
Option Compare Text

Function PjzFxa(Fxa) As VBProject
ThwIf_NotFxa Fxa, CSub
ThwIf_FfnNotExist Fxa, CSub, "Fxa"
Dim O As VBProject: Set O = PjzPjf(Xls.Vbe, Fxa): If Not IsNothing(O) Then Set PjzFxa = O: Exit Function
Set PjzFxa = OpnFx(Fxa).VBProject
End Function

Function HasFxa(Fxa$) As Boolean
HasFxa = HasEleS(PjfnAyV, Fn(Fxa))
End Function
Sub OpnFxa(Fxa$)
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
If HasFxa(Fxa) Then
    Inf CSub, "In Xls, there is Pjn = Fxa", "Fxa AllPj-In-Xls", Fxa, PjnyV
    Exit Sub
End If
Xls.Workbooks.Open Fxa
End Sub
Sub CrtFxa(Fxa$)  'Under Xls, crt an empty Fxa; set Pjn
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
If HasFxa(Fxa) Then Thw CSub, "In Xls, there is Pjn = Fxa", "Fxa AllPj-In-Xls", Fxa, PjnyV
Dim WB As Workbook: Set WB = Xls.Workbooks.Add
WB.SaveAs Fxa, XlFileFormat.xlOpenXMLAddIn 'Must save first, otherwise PjzFxa will fail.
PjzFxa(Fxa).Name = Fnn(RmvNxtNo(Fxa))
WB.Close True
End Sub

Function FrmFfny(Srcp$) As String()
Dim I
For Each I In Itr(Ffny(Srcp, "*.frm.txt"))
    PushI FrmFfny, I
Next
End Function



