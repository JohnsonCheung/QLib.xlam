Attribute VB_Name = "MApp_Rpt"
Option Explicit
Sub GenRpt(A As LidPm)
LnkImpzLidPm A
Run "GenOupTbl", A.Apn
CpyWszLidPm A, OupWbzInst(A.Apn), Vis:=True
End Sub

Private Sub ClsOupWbInst(Apn$)
Dim Wb As Workbook, F$
F = Fn(OupFx(Apn))
For Each Wb In Xls.Workbooks
    If IsInstFfn(Wb.FullName) And Wb.Name = F Then
        Wb.Close False
    End If
Next
End Sub

Private Function OupWbzInst(Apn$) As Workbook
ClsOupWbInst Apn
Dim OFx$
    OFx = OupFxInst(Apn)
    ExpTp Apn, OFx
Dim OWb As Workbook
    Set OWb = WbzFx(OFx)
    RfhWb(OWb, WFb(Apn)).Save
    Set OupWbzInst = OWb
End Function

Sub CpyWszLidPm(A As LidPm, ToOupWb As Workbook, Optional Vis As Boolean)
If Vis Then WbVis ToOupWb
End Sub

