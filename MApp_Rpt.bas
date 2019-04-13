Attribute VB_Name = "MApp_Rpt"
Option Explicit
Function OupFxzLidPm$(A As LidPm) 'Gen&Vis OupFx using LidPm as NxtFfn.
CpyFilzIfDif SyzOyPrp(A.Fil, "Ffn"), WPth(A.Apn)
LnkImpzLidPm A
Run "GenOupTbl", A.Apn
Dim OupWb As Workbook: Set OupWb = OupWbzNxt(A.Apn)
CpyWszLidPm A, OupWb, Vis:=True
FmtLozStdWb OupWb
RunAvzIgnEr "FmtOupWb", Av(OupWb)
OupWb.Save
OupFxzLidPm = OupWb.FullName
End Function

Private Sub ClsOupWb(Apn$)
Dim Wb As Workbook, F$
F = Fn(OupFx(Apn))
For Each Wb In Xls.Workbooks
    If IsNxtFfn(Wb.FullName) And Wb.Name = F Then
        Wb.Close False
    End If
Next
End Sub

Private Function OupWbzNxt(Apn$) As Workbook
ClsOupWb Apn
Dim OFx$
    OFx = OupFxzNxt(Apn)
    ExpTp Apn, OFx
Dim OWb As Workbook
    Set OWb = WbzFx(OFx)
    RfhWb(OWb, WFb(Apn)).Save
    Set OupWbzNxt = OWb
End Function

Sub CpyWszLidPm(A As LidPm, ToOupWb As Workbook, Optional Vis As Boolean)
If Vis Then WbVis ToOupWb
End Sub

