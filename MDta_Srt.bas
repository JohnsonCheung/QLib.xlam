Attribute VB_Name = "MDta_Srt"
Option Explicit

Function DrsSrt(A As Drs, Optional SrtByFF = "", Optional IsDes As Boolean) As Drs
Dim Fny$(): If SrtByFF = "" Then Fny = Sy(A.Fny()(0)) Else Fny = NyzNN(SrtByFF)
Set DrsSrt = Drs(A.Fny, DrySrt(A.Dry, IxAy(A.Fny, Fny), IsDes))
End Function

Function DrySrt(Dry(), Optional CC = 0, Optional IsDes As Boolean) As Variant()
If IsArray(CC) Then DrySrt = DrySrtzColIxAy(Dry, CC): Exit Function
DrySrt = DrySrtzCol(Dry, CC, IsDes)
End Function

Function DrySrtzCol(Dry(), ColIx, Optional IsDes As Boolean) As Variant()
Dim Col: Col = ColzDry(Dry, ColIx)
Dim Ix&(): Ix = IxAyzAySrt(Col, IsDes)
Dim J%
For J = 0 To UB(Ix)
   Push DrySrtzCol, Dry(Ix(J))
Next
End Function
Function DrySrtzColIxAy(Dry(), ColIxAy, Optional IsDes As Boolean) As Variant()
Stop
End Function

Function DtSrt(A As Dt, FF, Optional IsDes As Boolean) As Dt
Set DtSrt = Dt(A.DtNm, A.Fny, DrsSrt(DrszDt(A), FF, IsDes).Dry)
End Function

