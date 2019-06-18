Attribute VB_Name = "QVb_Dta_Ayab"
Option Explicit
Option Compare Text
Type Ayabc
    A As Variant
    B As Variant
    C As Variant
End Type
Type Ayab
    A As Variant
    B As Variant
End Type
Function Ayab(A, B) As Ayab
ThwIf_NotAy A, CSub
ThwIf_NotAy B, CSub
With Ayab
    .A = A
    .B = B
End With
End Function

Function Ayabc(A, B, C) As Ayabc
ThwIf_NotAy A, CSub
ThwIf_NotAy B, CSub
ThwIf_NotAy C, CSub
With Ayabc
    .A = A
    .B = B
    .C = C
End With
End Function

Function AyabzAyPfx(Ay, Pfx$) As Ayab
Dim O As Ayab
O.A = AyzReSi(Ay)
O.B = O.A
Dim S$, I
For Each I In Itr(Ay)
    S = I
    If HasPfx(S, Pfx) Then
        PushI O.B, S
    Else
        PushI O.A, S
    End If
Next
AyabzAyPfx = O
End Function

Function AyabzAyN(Ay, N&) As Ayab
AyabzAyN = Ayab(FstNEle(Ay, N), AyeFstNEle(Ay, N))
End Function

Function AyabczAyFE(Ay, FmIx&, EIx&) As Ayabc
Dim O As Ayabc
AyabczAyFE = Ayabc( _
    AywFE(Ay, 0, FmIx), _
    AywFE(Ay, FmIx, EIx), _
    AywFmIx(Ay, EIx))
End Function

Function AyabczAyFei(Ay, B As Fei) As Ayabc
AyabczAyFei = AyabczAyFE(Ay, B.FmIx, B.EIx)
End Function


Function DryzAyab(A, B) As Variant()
Dim J&
For J = 0 To Min(UB(A), UB(B))
    PushI DryzAyab, Array(A(J), B(J))
Next
End Function
Function DrszAyab(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2") As Drs
DrszAyab = Drs(Sy(N1, N2), DryzAyab(A, B))
End Function

