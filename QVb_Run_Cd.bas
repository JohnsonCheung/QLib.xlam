Attribute VB_Name = "QVb_Run_Cd"
Option Explicit
Private Const CMod$ = "MVb_Run_Cd."
Private Const Asm$ = "QVb"


Sub RunCdLy(CdLy$())
RunCd JnCrLf(CdLy)
End Sub

Sub RunCd(CdLines$)
Dim N$: N = "ZZZ_" & TmpNm
AddMthzCd N, CdLines
Run N
End Sub

Private Function RunCdMd() As CodeModule
EnsMd "ZTmpModForRun"
End Function
Private Sub AddMthzCd(MthNm$, CdLines$)
RunCdMd.AddFromString MthLines(MthNm$, CdLines)
End Sub
Private Function MthLines$(MthNm$, CdLines$)
Dim Lines$, L1$, L2$
L1 = "Sub ZZZ_" & MthNm & "()"
L2 = "End Sub"
MthLines = L1 & vbCrLf & CdLines & vbCrLf & L2
End Function

Private Property Get ZZCdLines$()
ZZCdLines = "MsgBox Now"
End Property


Sub TimFun(FunNN)
Dim B!, E!, F
For Each F In TermSy(FunNN)
    B = Timer
    Run F
    E = Timer
    Debug.Print F, "<-- Run"; E - B
Next
End Sub

Private Sub ZZ_TimFun()
TimFun "ZZA ZZB"
End Sub

Private Sub ZZA()
Dim J&, I&
For J = 0 To 100
    For I = 0 To 100
        Debug.Print I
    Next
Next
End Sub
Private Sub ZZB()
Dim J&, I&
For J = 0 To 100
    For I = 0 To 100
        Debug.Print I
    Next
Next
End Sub

