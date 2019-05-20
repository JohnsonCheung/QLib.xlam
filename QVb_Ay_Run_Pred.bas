Attribute VB_Name = "QVb_Ay_Run_Pred"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Run_Pred."
Private Const Asm$ = "QVb"

Function IsAllTruezItrPred_AyPred(A, Pred$) As Boolean
If Si(A) = 0 Then Exit Function
IsAllTruezItrPred_AyPred = IsAllTruezItrPred(A, Pred)
End Function

Function IsSomeTruezAyPred(A, Pred$) As Boolean
If Si(A) = 0 Then Exit Function
IsSomeTruezAyPred = IsSomFalsezItrPred(A, Pred)
End Function

Sub AyPredSplitAsg(A, Pred$, OTrueAy, OFalseAy)
Dim O1, O2
O1 = Resi(A)
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

