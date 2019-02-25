Attribute VB_Name = "MVb_Dic_SyDic"
Option Explicit
Sub PushItmzSyDic(A As Dictionary, K, Itm)
Dim M$()
If A.Exists(K) Then
    M = A(K)
    PushI M, Itm
    A(K) = M
Else
    A.Add K, Sy(Itm)
End If
End Sub

Sub ThwNotSyDic(A As Dictionary, Fun$)
If Not IsDiczSy(A) Then Thw Fun, "Given dictionary is not SyDic, all key is string and val is Sy", "Give-Dictionary", FmtDic(A)
End Sub

Function KeyToLikAyDic_T1LikssLy(TLikssLy$()) As Dictionary
Dim O As Dictionary
    Set O = Dic(TLikssLy)
Dim K
For Each K In O.Keys
    O(K) = SySsl(O(K))
Next
Set KeyToLikAyDic_T1LikssLy = O
End Function

