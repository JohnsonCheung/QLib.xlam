Attribute VB_Name = "QXls_Lo_Fmtr_Fmt"
Option Explicit
Private Const CMod$ = "MXls_Lo_Fmtr_Fmt."
Private Const Asm$ = "QXls"
Sub BrwSampLof()
Brw FmtLof(SampLof)
End Sub
Function FmtLof(Lof$()) As String()
FmtLof = FmtSpec(Lof, LofT1nn, 2)
End Function

Function FmtSpec(Spec$(), Optional T1nn$, Optional FmtFstNTerm% = 1) As String()
Dim mT1Ay$()
    If IsMissing(T1nn) Then
        mT1Ay = T1Sy(Spec)
    Else
        mT1Ay = TermSy(T1nn)
    End If
Dim O$()
    Dim T$, I
    For Each I In mT1Ay
        T = I
        PushIAy O, AywT1(Spec, T)
    Next
    Dim M$(): M = SyeT1Sy(Spec, mT1Ay)
    If Si(M) > 0 Then
        PushI O, FmtQQ("# Error: in not T1Sy(?)", TLin(mT1Ay))
        PushIAy O, M
    End If
FmtSpec = FmtAyNTerm(O, FmtFstNTerm)
End Function
