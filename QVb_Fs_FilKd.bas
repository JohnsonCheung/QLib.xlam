Attribute VB_Name = "QVb_Fs_FilKd"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Fs_Ffn_MisEr."
Type FilKd
    Kd As String
    Ffn As String
End Type
Type FilKds: N As Byte: Ay() As FilKd: End Type
Function KdSyzFilKds(A As FilKds) As String()
Dim J%
For J = 0 To A.N - 1
    PushI KdSyzFilKds, A.Ay(J).Kd
Next
End Function

Function FfnSyzFilKds(A As FilKds) As String()
Dim J%
For J = 0 To A.N - 1
    PushI FfnSyzFilKds, A.Ay(J).Ffn
Next
End Function

Function FilKdszLy(Ly$()) As FilKds
Dim I, Lin$
For Each I In Itr(Ly)
    Lin = I
    PushFilKd FilKdszLy, FilKdzLin(Lin)
Next
End Function
Function FilKdzLin(Lin$) As FilKd
With BrkSpc(Lin)
FilKdzLin = FilKd(.S1, .S2)
End With
End Function

Sub ThwIfMisFilKds(A As FilKds, Fun$)
ThwIfEr MsgzMisFilKds(FilKdszWhMis(A)), Fun
End Sub

Function FilKd(Kd$, Ffn$) As FilKd
FilKd.Kd = Kd
FilKd.Ffn = Ffn
End Function
Function SngFilKd(A As FilKd) As FilKds
PushFilKd SngFilKd, A
End Function

Sub PushFilKd(O As FilKds, M As FilKd)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function MsgzMisFilKd(Mis As FilKd) As String()
MsgzMisFilKd = MsgzMisFilKds(SngFilKd(Mis))
End Function

Function MsgzMisFilKds(Mis As FilKds) As String()
If Mis.N = 0 Then Exit Function
Dim M$, Ny$(), Av(), NN$
M = FmtQQ("? file not found", Mis.N)
Dim J%
For J = 0 To Mis.N - 1
    With Mis.Ay(J)
    PushI Ny, "In Path":        PushI Av, Pth(.Ffn)
    PushI Ny, "Missing " & .Kd: PushI Av, Fn(.Ffn)
    End With
Next
NN = JnSpc(QuoteSqBktIfzSy(Ny))
MsgzMisFilKds = LyzMsgNap(M, NN, Av)
End Function

Function FilKdszWhMis(A As FilKds) As FilKds
Dim J%
For J = 0 To A.N - 1
    With A.Ay(J)
        If Not HasFfn(.Ffn) Then PushFilKd FilKdszWhMis, A.Ay(J)
    End With
Next
End Function
