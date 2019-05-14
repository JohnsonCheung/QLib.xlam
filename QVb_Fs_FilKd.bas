Attribute VB_Name = "QVb_Fs_FilKd"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Fs_Ffn_MisEr."
Type KdFil
    Kd As String
    Ffn As String
End Type
Type KdFils: N As Byte: Ay() As KdFil: End Type
Function KdAyzKdFils(A As KdFils) As String()
Dim J%
For J = 0 To A.N - 1
    PushI KdAyzKdFils, A.Ay(J).Kd
Next
End Function

Function FfnyzKdFils(A As KdFils) As String()
Dim J%
For J = 0 To A.N - 1
    PushI FfnyzKdFils, A.Ay(J).Ffn
Next
End Function
Function KdFilszVbl(Vbl$) As KdFils
KdFilszVbl = KdFils(LyzVbl(Vbl))
End Function

Function KdFils(Ly$()) As KdFils
Dim Lin
For Each Lin In Itr(Ly)
    PushKdFil KdFils, KdFilzLin(Lin)
Next
End Function
Function KdFilzLin(Lin) As KdFil
With BrkSpc(Lin)
KdFilzLin = KdFil(.S1, .S2)
End With
End Function

Sub ThwIf_MisKdFils(A As KdFils, Fun$)
ThwIf_Er MsgzMisKdFils(KdFilszWhMis(A)), Fun
End Sub

Function KdFil(Kd, Ffn) As KdFil
KdFil.Kd = Kd
KdFil.Ffn = Ffn
End Function
Function SngKdFil(A As KdFil) As KdFils
PushKdFil SngKdFil, A
End Function

Sub PushKdFil(O As KdFils, M As KdFil)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Function MsgzMisKdFil(Mis As KdFil) As String()
MsgzMisKdFil = MsgzMisKdFils(SngKdFil(Mis))
End Function

Function MsgzMisKdFils(Mis As KdFils) As String()
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
MsgzMisKdFils = LyzMsgNap(M, NN, Av)
End Function

Function KdFilszWhMis(A As KdFils) As KdFils
Dim J%
For J = 0 To A.N - 1
    With A.Ay(J)
        If Not HasFfn(.Ffn) Then PushKdFil KdFilszWhMis, A.Ay(J)
    End With
Next
End Function
