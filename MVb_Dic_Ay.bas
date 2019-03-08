Attribute VB_Name = "MVb_Dic_Ay"
Option Explicit
Function DicAdd(A As Dictionary, B As Dictionary) As Dictionary
Dim O As Dictionary
Set O = DicClone(A)
PushDic O, B
Set DicAdd = O
End Function

Function DicAyAdd(DicAy() As Dictionary) As Dictionary
'WarnLin if Key Has
Dim I
For Each I In DicAy
    'PushDic DicMgeAy, CvDic(I)
Next
End Function

Function ColDicAyKey(DicAy() As Dictionary, Key) As Variant()
Dim U%: U = UB(DicAy)
Dim O()
Dim I, Dic As Dictionary, J%
J = 1
'O(0) = K
For Each I In DicAy
   Set Dic = I
'   If Dic.Exists(K) Then O(J) = Dic(K)
Next
'DicAyDr = O
End Function
