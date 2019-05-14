Attribute VB_Name = "QVb_Dic_Dice"
Option Explicit
Private Const CMod$ = "MVb_Dic_Ay."
Private Const Asm$ = "QVb"
Function DicAyzAp(ParamArray DicAp()) As Dictionary()
Dim Av(): Av = DicAp: If Si(Av) = 0 Then Exit Function
Dim I
For Each I In Av
    If Not IsDic(I) Then Thw CSub, "Some itm is not Dic", "TypeName-Ay", TyNyzAy(Av)
    PushObj DicAyzAp, CvDic(I)
Next
End Function

Function DiceKeySet(A As Dictionary, ExlKeySet As Aset) As Dictionary
Dim K
Set DiceKeySet = New Dictionary
For Each K In A.Keys
    If Not ExlKeySet.Has(K) Then
        DiceKeySet.Add K, A(K)
    End If
Next
End Function

