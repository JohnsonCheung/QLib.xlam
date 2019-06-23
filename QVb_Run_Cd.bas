Attribute VB_Name = "QVb_Run_Cd"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Run_Cd."
Private Const Asm$ = "QVb"


Sub RunCdLy(CdLy$())
RunCd JnCrLf(CdLy)
End Sub

Sub RunCd(CdLines$)
Dim N$: N = "Z_" & TmpNm
AddMthzCd N, CdLines
Run N
End Sub

Private Function RunCdMd() As CodeModule
'EnsMd "ZTmpModForRun"
End Function
Private Sub AddMthzCd(Mthn, CdLines$)
RunCdMd.AddFromString MthL(Mthn, CdLines)
End Sub
Private Function MthL$(Mthn, CdLines$)
Dim Lines$, L1$, L2$
L1 = "Private Sub Z_" & Mthn & "()"
L2 = "End Sub"
MthL = L1 & vbCrLf & CdLines & vbCrLf & L2
End Function

Private Function Y_CdLines$()
Y_CdLines = "MsgBox Now"
End Function


Sub TimFun(FunNN)
Dim B!, E!, F
For Each F In TermAy(FunNN)
    B = Timer
    Run F
    E = Timer
    Debug.Print F, "<-- Run"; E - B
Next
End Sub

Private Sub Z_TimFun()
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

