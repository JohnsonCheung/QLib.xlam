Attribute VB_Name = "QVb_Dta_Ayab"
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
Function Ayabc(A, B, C) As Ayab
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
Dim O As New Ayab
O.A = Resi(Ay)
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

Function AyabByN(Ay, N&) As Ayab
Set AyabByN = Ayab(AywFstNEle(Ay, N), AyeFstNEle(Ay, N))
End Function

Function AyabByEle(Ay, Ele) As Ayab
Dim O As Ayab
O.A = Resi(Ay)
O.B = O.A
Dim J%
For J = 0 To UB(Ay)
    If Ay(J) = Ele Then Exit For
    PushI O.A, Ay(J)
Next
For J = J + 1 To UB(Ay)
    PushI O.B, Ay(J)
Next
AyabByEle = O
End Function

Function AyabcByFmTo(Ay, FmIx&, EIx&) As Ayabc
Dim O As New Ayabc
Set AyabcByFmTo = O.Init( _
    AywFT(Ay, 0, FmIx - 1), _
    AywFT(Ay, FmIx, EIx), _
    AywFmIx(Ay, EIx + 1))
End Function

Function AyabcByFEIx(Ay, B As FEIx) As Ayabc
Set AyabcByFEIx = AyabcByFmTo(Ay, B.FmIx, B.EIx)
End Function


