Attribute VB_Name = "MVb_Dic_Has"
Option Explicit

Function DicHasAllKeyIsNm(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
DicHasAllKeyIsNm = True
End Function

Function DicHasAllValIsStr(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsStr(A(K)) Then Exit Function
Next
DicHasAllValIsStr = True
End Function

Function DicHasBlankKey(A As Dictionary) As Boolean
If A.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
   If Trim(K) = "" Then DicHasBlankKey = True: Exit Function
Next
End Function

Function DicHasK(A As Dictionary, K$) As Boolean
DicHasK = A.Exists(K)
End Function

Function DicHasKeyLvs(A As Dictionary, KeyLvs) As Boolean
DicHasKeyLvs = DicHasKy(A, SySsl(KeyLvs))
End Function

Sub DicHasKeyssAss(A As Dictionary, KeySS$)
DicHasKyAss A, SySsl(KeySS)
End Sub

Function DicHasKeySsl(A As Dictionary, KeySsl) As Boolean
DicHasKeySsl = A.Exists(SySsl(KeySsl))
End Function

Function DicHasKy(A As Dictionary, Ky) As Boolean
Ass IsArray(Ky)
If Si(Ky) = 0 Then Stop
Dim K
For Each K In Ky
   If Not A.Exists(K) Then
       Debug.Print FmtQQ("Dix.HasKy: Key(?) is Missing", K)
       Exit Function
   End If
Next
DicHasKy = True
End Function

Sub DicHasKyAss(A As Dictionary, Ky)
Dim K
For Each K In Ky
   If Not A.Exists(K) Then Debug.Print K: Stop
Next
End Sub

Function DicKeysIsAllStr(A As Dictionary) As Boolean
DicKeysIsAllStr = IsAllStrAy(A.Keys)
End Function

Private Sub Z_DicKeysIsAllStr()
Dim A As Dictionary
GoSub T1
Exit Sub
T1:
    Set A = New Dictionary
    Dim J&
    For J = 1 To 10000
        A.Add CStr(J), J
    Next
    Ept = True
    GoSub Tst
    '
    A.Add 10001, "X"
    Ept = False
    GoTo Tst
Tst:
    Act = DicKeysIsAllStr(A)
    C
    Return
End Sub

