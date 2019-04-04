Attribute VB_Name = "MVb_Ay_Op_Brk"
Option Explicit
Function AyabzAyPfx(Ay, Pfx) As AyAB
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
AyabzAyPfx = O
End Function
Function AyabzAyEle(Ay, Ele) As AyAB
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
AyabzAyEle = O
End Function

Function AyabCzFT(Ay, FmIx&, ToIx&) As AyABC
Dim O As New AyABC
Set AyabCzFT = O.Init( _
    AywFT(Ay, 0, FmIx - 1), _
    AywFT(Ay, FmIx, ToIx), _
    AywFmIx(Ay, ToIx + 1))
End Function

Function AyabCzFTIx(Ay, B As FTIx) As AyABC
Set AyabCzFTIx = AyabCzFT(Ay, B.FmIx, B.ToIx)
End Function
