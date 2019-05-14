Attribute VB_Name = "QIde_Mth_TopRmk"
Option Explicit
Private Const CMod$ = "MIde_Mth_TopRmk."
Private Const Asm$ = "QIde"
Private Sub Z_MthFEIxszSrcMth()
Dim Src$(), Mthn, WiTopRmk As Boolean
Dim Ept As FEIxs, Act As FEIxs

Src = SrczMdn("IdeMthFEIx")
PushFEIx Ept, FEIx(2, 11)
GoSub Tst

Exit Sub
Tst:
    Act = MthFEIxszSN(Src, Mthn, WiTopRmk)
    If Not IsEqFEIxs(Act, Ept) Then Stop
    Return
End Sub

Function RmvBlankLin(Ay) As String()
Dim I
For Each I In Itr(Ay)
    PushNonBlank RmvBlankLin, I
Next
End Function

Function TopRmkLines$(Src$(), MthIx)
TopRmkLines = JnCrLf(TopRmkLy(Src, MthIx))
End Function
Function TopRmkLyzSIW(Src$(), MthIx, WiTopRmk As Boolean) As String()
If Not WiTopRmk Then Exit Function
TopRmkLyzSIW = TopRmkLy(Src, MthIx)
End Function
Function TopRmkLy(Src$(), MthIx) As String()
Dim Fm&: Fm = TopRmkIx(Src, MthIx): If Fm = -1 Then Exit Function
TopRmkLy = RmvBlankLin(AywFT(Src, Fm, MthIx - 1))
End Function

Function TopRmkIx&(Src$(), MthIx)
If MthIx <= 0 Then Exit Function
Dim J&, L$
TopRmkIx = MthIx
For J = MthIx - 1 To 0 Step -1
    L = LTrim(Src(J))
    Select Case True
    Case L = ""
    Case FstChr(L) = "'": TopRmkIx = J
    Case Else: Exit Function
    End Select
Next
End Function

Function TopRmkLno(Md As CodeModule, MthLno)
Dim J&, L$
TopRmkLno = MthLno
If MthLno = 0 Then Exit Function
For J = MthLno - 1 To 1 Step -1
    L = LTrim(Md.Lines(J, 1))
    Select Case True
    Case L = ""
    Case FstChr(L) = "'": TopRmkLno = J
    Case Else: Exit Function
    End Select
Next
End Function
