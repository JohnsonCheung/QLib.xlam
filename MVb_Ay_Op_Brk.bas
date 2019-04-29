Attribute VB_Name = "MVb_Ay_Op_Brk"
Option Explicit
Function AyabByPfx(Ay, Pfx$) As Ayab
Dim O As New Ayab
O.A = AyCln(Ay)
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
AyabByPfx = O
End Function
Function Ayab(A, B) As Ayab
Dim O As New Ayab
Set Ayab = O.Init(A, B)
End Function

Function AyabByN(Ay, N) As Ayab
Set AyabByN = Ayab(AywFstNEle(Ay, N), AyeFstNEle(Ay, N))
End Function

Function AyabByEle(Ay, Ele) As Ayab
Dim O As Ayab
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
AyabByEle = O
End Function

Function AyabcByFmTo(Ay, FmIx&, ToIx&) As AyABC
Dim O As New AyABC
Set AyabcByFmTo = O.Init( _
    AywFT(Ay, 0, FmIx - 1), _
    AywFT(Ay, FmIx, ToIx), _
    AywFmIx(Ay, ToIx + 1))
End Function

Function AyabcByFTIx(Ay, B As FTIx) As AyABC
Set AyabcByFTIx = AyabcByFmTo(Ay, B.FmIx, B.ToIx)
End Function
