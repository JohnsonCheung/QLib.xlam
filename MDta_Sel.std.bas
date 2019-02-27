Attribute VB_Name = "MDta_Sel"
Option Explicit

Function DrySel(A, IxAy) As Variant()
Dim Dr
For Each Dr In Itr(A)
    PushI DrySel, AywIxAy(Dr, IxAy)
Next
End Function

Function DrySelIxAp(A, ParamArray IxAp()) As Variant()
Dim IxAy(): IxAy = IxAp
DrySelIxAp = DrySel(A, IxAy)
End Function

Function DrsSel(A As DRs, FF) As DRs
Dim Fny$(): Fny = CvNy(FF)
If IsEqAy(A.Fny, Fny) Then Set DrsSel = A: Exit Function
ThwNotSuperAy A.Fny, Fny
Set DrsSel = DRs(Fny, DrySel(A.Dry, IxAy(A.Fny, Fny)))
End Function

Private Sub Z_DrsSel()
'BrwDrs DrsSel(Vmd.MthDrs, "MthNm Mdy Ty MdNm")
'BrwDrs Vmd.MthDrs
End Sub

Function DtSel(A As DT, FF) As DT
Set DtSel = DtzDrs(DrsSel(DrszDt(A), FF), A.DtNm)
End Function


Private Sub Z()
Z_DrsSel
MDta_Sel:
End Sub
