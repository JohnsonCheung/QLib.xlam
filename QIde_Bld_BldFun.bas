Attribute VB_Name = "QIde_Bld_BldFun"

Function PjzFxa(Fxa) As VBProject
ThwIf_NotFxa Fxa, CSub
ThwIf_FfnNotExist Fxa, CSub, "Fxa"
Dim O As VBProject: Set O = PjzPjf(Xls.Vbe, Fxa): If Not IsNothing(O) Then Set PjzFxa = O: Exit Function
Set PjzFxa = OpnFx(Fxa).VBProject
End Function

Function HasFxa(Fxa$) As Boolean
HasFxa = HasEleS(PjFnAyC, Fn(Fxa))
End Function

Sub CrtFxa(Fxa$)  'Under Xls, crt an empty Fxa; set Pjn
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
If HasFxa(Fxa) Then Thw CSub, "In Xls, there is Pjn = Fxa", "Fxa AllPj-In-Xls", Fxa, Pjny
Dim Wb As Workbook: Set Wb = Xls.Workbooks.Add
Wb.SaveAs Fxa, XlFileFormat.xlOpenXMLAddIn 'Must save first, otherwise PjzFxa will fail.
PjzFxa(Fxa).Name = Fnn(RmvNxtNo(Fxa))
Wb.Close True
End Sub

Function FrmFfny(Srcp$) As String()
Dim I
For Each I In Itr(Ffny(Srcp, "*.frm.txt"))
    PushI FrmFfny, I
Next
End Function


