Attribute VB_Name = "MIde_Mth_TopRmk"
Option Explicit
Private Sub Z_MthFTIxAyzSrcMth()
Dim Src$(), MthNm, WiTopRmk As Boolean
Dim Ept() As FTIx, Act() As FTIx

Src = SrczMdNm("IdeMthFTIx")
PushObj Ept, FTIx(2, 11)
GoSub Tst

Exit Sub
Tst:
    Act = MthFTIxAyzSrcMth(Src, MthNm, WiTopRmk)
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
MthTopRmkLy = AyRmvBlankLin(AywFT(Src, Fm, MthFmIx - 1))
End Function

Function MthTopRmkIx&(Src$(), MthFmIx)
Dim J&, L
MthTopRmkIx = MthFmIx
If MthFmIx = 0 Then Exit Function
For J = MthFmIx - 1 To 0 Step -1
    L = LTrim(Src(J))
    Select Case True
    Case L = ""
    Case FstChr(L) = "'": MthTopRmkIx = J
    Case Else: Exit Function
    End Select
Next
End Function

Function MthTopRmkLno&(Md As CodeModule, MthLno)
Dim J&, L
MthTopRmkLno = MthLno
If MthLno = 0 Then Exit Function
For J = MthLno - 1 To 1 Step -1
    L = LTrim(Md.Lines(J, 1))
    Select Case True
    Case L = ""
    Case FstChr(L) = "'": MthTopRmkLno = J
    Case Else: Exit Function
    End Select
Next
End Function
