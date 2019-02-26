Attribute VB_Name = "MDta_Srt"
Option Explicit

Function SrtDrs(A As DRs, FF, Optional Des As Boolean) As DRs
'Set DrsSrt = Drs(A.Fny, SrtDry(A.Dry, IxAy(A.Fny, FF), IsDes))
End Function

Function SrtDry(A(), Optional ColIxAy, Optional IsDes As Boolean) As Variant()
Dim Col: Col = ColzDry(A, ColIxAy)
Dim Ix&(): Ix = AySrtIntoIxAy(Col, IsDes)
Dim J%, O()
For J = 0 To UB(Ix)
   Push O, A(Ix(J))
Next
SrtDry = O
End Function

Function SrtDt(A As Dt, FF, Optional IsDes As Boolean) As Dt
'Set Dt_Srt = Dt(A.DtNm, A.Fny, DrsSrt(DrszDt(A), FF, IsDes).Dry)
End Function

