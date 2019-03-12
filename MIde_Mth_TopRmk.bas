Attribute VB_Name = "MIde_Mth_TopRmk"
Option Explicit
Private Sub Z_MthWTopRmkMthFTixAyzSrcMth()
Dim Src$(), MthNm, WithTopRmk As Boolean
Dim Ept() As FTIx, Act() As FTIx

Src = SrcMdNm("IdeMthFTIx")
PushObj Ept, FTIx(2, 11)
GoSub Tst

Exit Sub
Tst:
    Act = MthFTIxAyzSrcMth(Src, MthNm, WithTopRmk)
    If Not IsEqFTIxAy(Act, Ept) Then Stop
    Return
End Sub

Function AyRmvBlankLin(Ay) As String()
Dim L
For Each L In Itr(Ay)
    PushNonBlankStr AyRmvBlankLin, L
Next
End Function

Function MthTopRmkLy(Src$(), MthFmIx) As String()
Dim Fm&: Fm = MthTopRmkIx(Src, MthFmIx): If Fm = -1 Then Exit Function
MthTopRmkLy = AyRmvBlankLin(AywFT(Src, Fm, MthFmIx + 1))
End Function

Function MthTopRmkIx&(Src$(), MthFmIx)
Dim J&, L
If IsCdLin(Src(MthFmIx)) Then
    MthTopRmkIx = -1
    Exit Function
End If

For J = MthFmIx - 1 To 0 Step -1
    L = Src(J)
    If IsCdLin(L) Then
        MthTopRmkIx = J + 1
        Exit Function
    End If
Next
MthTopRmkIx = -1
End Function
Function MthTopRmkLnoMdFm&(Md As CodeModule, MthLno)
Dim J&, L
For J = MthLno To 1 Step -1
    L = Md.Lines(J, 1)
    If IsCdLin(L) Then
        MthTopRmkLnoMdFm = J - 1
        Exit Function
    End If
Next
End Function

