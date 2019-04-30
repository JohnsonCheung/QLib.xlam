Attribute VB_Name = "MDta_Sel"
Option Explicit

Function DrySel(Dry(), IxAy&()) As Variant()
Dim Drv
For Each Drv In Itr(Dry)
    PushI DrySel, AywIxAy(Drv, IxAy)
Next
End Function

Function DrsSel(A As Drs, FF$) As Drs
Dim Fny$(): Fny = TermAy(FF)
If IsEqAy(A.Fny, Fny) Then DrsSel = A: Exit Function
ThwNotSuperAy A.Fny, Fny
DrsSel = Drs(Fny, DrySel(A.Dry, IxAy(A.Fny, Fny)))
End Function

Private Sub Z_DrsSel()
'BrwDrs DrsSel(Vmd.MthDrs, "MthNm Mdy Ty MdNm")
'BrwDrs Vmd.MthDrs
End Sub

Function DtSel(A As Dt, FF$) As Dt
DtSel = DtzDrs(DrsSel(DrszDt(A), FF), A.DtNm)
End Function


Private Sub Z()
Z_DrsSel
MDta_Sel:
End Sub
