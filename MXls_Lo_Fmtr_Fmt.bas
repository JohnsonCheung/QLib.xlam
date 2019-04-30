Attribute VB_Name = "MXls_Lo_Fmtr_Fmt"
Option Explicit
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
        mT1Ay = TermAy(T1nn)
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
