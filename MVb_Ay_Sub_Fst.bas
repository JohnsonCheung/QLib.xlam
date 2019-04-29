Attribute VB_Name = "MVb_Ay_Sub_Fst"
Option Explicit
Function ShfFstEle(OAy)
ShfFstEle = FstEle(OAy)
OAy = AyeFstNEle(OAy)
End Function

Function FstEle(Ay)
If Si(Ay) = 0 Then Exit Function
Asg Ay(0), FstEle
End Function

Function FstEleEv(Ay, V)
If HasEle(Ay, V) Then FstEleEv = V
End Function

Function FstEleLik$(A, Lik$)
If Si(A) = 0 Then Exit Function
Dim X
For Each X In A
    If X Like Lik Then FstEleLik = X: Exit Function
Next
End Function

Function FstElezPfxAy$(PfxAy$(), Lin$)
Dim I, P$
For Each I In PfxAy
    P = I
    If HasPfx(Lin, P) Then FstElezPfxAy = P: Exit Function
Next
End Function
Function FstEleInAset(Ay, InAset As Aset)
Dim I
For Each I In Ay
    If InAset.Has(I) Then FstEleInAset = I: Exit Function
Next
End Function

Function FstElePredPX(A, PX$, P)
If Si(A) = 0 Then Exit Function
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
If Si(A) = 0 Then Exit Function
Dim X
For Each X In Itr(A)
    If Run(XP, X, P) Then Asg X, FstElePredXP: Exit Function
Next
End Function

Function FstElezRmvT1$(Sy$(), T1$)
FstElezRmvT1 = RmvT1(FstElezT1(Sy, T1))
End Function

Function FstElezT1$(Sy$(), T1$)
Dim I, S$
For Each I In Itr(Sy)
    S = I
    If HasT1(S, T1) Then FstElezT1 = S: Exit Function
Next
End Function

Function FstElezT2$(Sy$(), T2$)
Dim I, S$
For Each I In Itr(Sy)
    S = I
    If HasT2(S, T2) Then FstElezT2 = S: Exit Function
Next
End Function

Function FstElezTT$(Sy$(), T1$, T2$)
Dim I, S$
For Each I In Itr(Sy)
    S = I
    If HasTT(S, T1, T2) Then FstElezTT = S: Exit Function
Next
End Function


Function FstEleRmvTT$(Sy$(), T1$, T2$)
Dim I, L$, X1$, X2$, Rst$
For Each I In Itr(Sy)
    L = I
    AsgN2tRst L, X1, X2, Rst
    If X1 = T1 Then
        If X2 = T2 Then
            FstEleRmvTT = L
            Exit Function
        End If
    End If
Next
End Function


