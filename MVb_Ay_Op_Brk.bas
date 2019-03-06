Attribute VB_Name = "MVb_Ay_Op_Brk"
Option Explicit
Function AyABzAyPfx(Ay, Pfx) As AyAB
Dim O As New AyAB
O.A = AyCln(Ay)
O.B = O.A
Dim S
For Each S In Itr(Ay)
    If HasPfx(S, Pfx) Then
        PushI O.B, S
    Else
        PushI O.A, S
    End If
Next
AyABzAyPfx = O
End Function
Function AyABzAyEle(Ay, Ele) As AyAB
Dim O As AyAB
O.A = AyCln(Ay)
O.B = O.A
Dim J%
For J = 0 To UB(Ay)
    If Ay(J) = Ele Then Exit For
    PushI O.A, Ay(J)
Next
For J = J + 1 To UB(Ay)
    PushI O.B, Ay(J)
Next
AyABzAyEle = O
End Function

Function AyABCzAyFT(Ay, FmIx&, ToIx&) As AyABC
Dim O As AyABC
O.A = AywFT(Ay, 0, FmIx - 1)
O.B = AywFT(Ay, FmIx, ToIx)
O.C = AywFmIx(Ay, ToIx + 1)
AyABCzAyFT = O
End Function

Function AyABCzAyFTIx(A, B As FTIx) As AyABC
AyABCzAyFTIx = AyABCzAyFT(A, B.FmIx, B.ToIx)
End Function