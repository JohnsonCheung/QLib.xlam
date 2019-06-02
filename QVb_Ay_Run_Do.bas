Attribute VB_Name = "QVb_Ay_Run_Do"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Run_Do."
Private Const Asm$ = "QVb"
Sub DoAy(A, FunNm$)
If Si(A) = 0 Then Exit Sub
Dim I
For Each I In A
    Run FunNm, I
Next
End Sub

Sub DoAyABX(Ay, ABX$, A, B)
Dim X
For Each X In Itr(Ay)
    Run ABX, A, B, X
Next
End Sub

Sub DoAyAXB(Ay, AXB$, A, B)
Dim X
For Each X In Itr(Ay)
    Run AXB, A, X, B
Next
End Sub

Sub DoAyPPXP(A, PPXP$, P1, P2, P3)
Dim X
For Each X In Itr(A)
    Run PPXP, P1, P2, X, P3
Next
End Sub

Sub DoAyPX(A, PX$, P)
Dim X
For Each X In Itr(A)
    Run PX, P, X
Next
End Sub

Sub DoAyXP(A, Xp$, P)
Dim X
For Each X In Itr(A)
    Run Xp, X, P
Next
End Sub
