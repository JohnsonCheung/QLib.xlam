Attribute VB_Name = "QIde_Bld_BldFun"

Function PjzFxa(Fxa$) As VBProject
ThwIfNotFxa Fxa, CSub
ThwIfFfnNotExist Fxa, CSub, "Fxa"
Dim O As VBProject: Set O = PjzPjf(Xls.Vbe, Fxa): If Not IsNothing(O) Then Set PjzFxa = O: Exit Function
Set PjzFxa = OpnFx(Fxa).VBProject
End Function

Function HasFxa(Fxa$) As Boolean
HasFxa = HasEleS(PjFnSyC, Fn(Fxa))
End Function

Sub CrtFxa(Fxa$)  'Under Xls, crt an empty Fxa; set PjNm
If Not IsFxa(Fxa) Then Thw CSub, "Not a Fxa", "Fxa", Fxa
If HasFxa(Fxa) Then Thw CSub, "In Xls, there is PjNm = Fxa", "Fxa AllPj-In-Xls", Fxa, PjNy
Dim Wb As Workbook: Set Wb = Xls.Workbooks.Add
Wb.SaveAs Fxa, XlFileFormat.xlOpenXMLAddIn 'Must save first, otherwise PjzFxa will fail.
PjzFxa(Fxa).Name = Fnn(RmvNxtNo(Fxa))
Wb.Close True
End Sub

Function FrmFfnSy(Srcp$) As String()
Dim Ffn$, I
For Each I In FfnSy(Scrp)
    Ffn = I
    If HasSfx(Ffn, ".frm.txt") Then
        PushI FrmFfnSy, Ffn
    End If
Next
End Function


