Attribute VB_Name = "MVb_Ay_Sub_Fst"
Option Explicit
Function ShfFstEle(OAy)
ShfFstEle = FstEle(OAy)
OAy = AyeFstNEle(OAy)
End Function

Function FstEle(Ay)
If Sz(Ay) = 0 Then Exit Function
Asg Ay(0), FstEle
End Function

Function FstEleEv(A, V)
If HasEle(A, V) Then FstEleEv = V
End Function

Function FstEleLik$(A, Lik$)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In A
    If X Like Lik Then FstEleLik = X: Exit Function
Next
End Function

Function FstElePfx$(PfxAy, Lin$)
Dim P
For Each P In PfxAy
    If HasPfx(Lin, CStr(P)) Then FstElePfx = P: Exit Function
Next
End Function

Function FstElePredPX(A, PX$, P)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In Itr(A)
    If Run(PX, P, X) Then Asg X, FstElePredPX: Exit Function
Next
End Function

Function FstElePredXABTrue(Ay, XAB$, A, B)
Dim X
For Each X In Itr(Ay)
    If Run(XAB, X, A, B) Then Asg X, FstElePredXABTrue: Exit Function
Next
End Function

Function FstElePredXP(A, XP$, P)
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In Itr(A)
    If Run(XP, X, P) Then Asg X, FstElePredXP: Exit Function
Next
End Function

Function FstEleRmvT1$(Ay, T1Val, Optional IgnCas As Boolean)
FstEleRmvT1 = RmvT1(FstEleT1(Ay, T1Val, IgnCas))
End Function

Function FstEleT1$(Ay, T1Val, Optional IgnCas As Boolean)
Dim L
For Each L In Itr(Ay)
    If IsEqStr(T1(L), T1Val, IgnCas) Then FstEleT1 = L: Exit Function
Next
End Function

Function FstEleT2$(A, T2)
FstEleT2 = FstElePredXP(A, "HasL_T2", T2)
End Function

Function FstEleTT$(A, T1, T2)
FstEleTT = FstElePredXABTrue(A, "HasTT", T1, T2)
End Function


Function FstEleRmvTT$(A, T1$, T2$)
Dim X, X1$, X2$, Rst$
For Each X In Itr(A)
    Asg2TRst X, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            FstEleRmvTT = X
            Exit Function
        End If
    End If
Next
End Function


