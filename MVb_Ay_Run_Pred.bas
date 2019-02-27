Attribute VB_Name = "MVb_Ay_Run_Pred"
Option Explicit

Function IsAllTrue_ItrPred_AyPred(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
IsAllTrue_ItrPred_AyPred = IsAllTrue_ItrPred(A, Pred)
End Function

Function IsSomTrue_AyPred(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
IsSomTrue_AyPred = IsSomFalse_ItrPred(A, Pred)
End Function

Sub AyPredSplitAsg(A, Pred$, OTrueAy, OFalseAy)
Dim O1, O2
O1 = AyCln(A)
O2 = O1
Dim X
For Each X In Itr(A)
    If Run(Pred, X) Then
        Push OTrueAy, X
    Else
        Push OFalseAy, X
    End If
Next
End Sub


Function IsAllFalse_ItrPred_AyPred(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
IsAllFalse_ItrPred_AyPred = IsAllFalse_ItrPred(A, Pred)
End Function
