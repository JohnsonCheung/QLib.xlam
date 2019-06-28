Attribute VB_Name = "QVb_Dic_CntDic"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Dic_DiKqCnt."
Private Const Asm$ = "QVb"
Function FmtDiKqCnt(Ay, Optional Opt As EmCnt) As String()
FmtDiKqCnt = FmtS12s(SwapS12s(S12szDic(DiKqCnt(Ay, Opt))), N1:="Cnt", N2:="Mth")
End Function

Function CntzAyD(Ay, Optional C As VbCompareMethod = vbTextCompare) As Dictionary
'Ret : :DiKqCnt #Cnt-Ay-Dic ! Cnt-Ay-ret-as-DiKqCnt.  :DiKqCnt is val is a cnt (num).  %Cnt always :DiKqCnt
Dim O As New Dictionary, I
O.CompareMode = C
For Each I In Itr(Ay)
    If O.Exists(I) Then
        O(I) = O(I) + 1
    Else
        O.Add I, 1
    End If
Next
Set CntzAyD = O
End Function
Function DiKqCntwDup(DiKqCnt As Dictionary) As Dictionary
Set DiKqCntwDup = New Dictionary
Dim Cnt&, K
For Each K In DiKqCnt.Keys
    Cnt = DiKqCnt(K)
    If Cnt > 1 Then DiKqCntwDup.Add K, Cnt
Next
End Function

Function DiKqCntwSng(DiKqCnt As Dictionary) As Dictionary
Set DiKqCntwSng = New Dictionary
Dim Cnt&, K
For Each K In DiKqCnt.Keys
    Cnt = DiKqCnt(K)
    If Cnt = 1 Then DiKqCntwSng.Add K, Cnt
Next
End Function
Function DiKqCntwEmCnt(DiKqCnt As Dictionary, B As EmCnt) As Dictionary
Select Case B
Case EmCnt.EiCntDup: Set DiKqCntwEmCnt = DiKqCntwDup(DiKqCnt)
Case EmCnt.EiCntSng: Set DiKqCntwEmCnt = DiKqCntwSng(DiKqCnt)
Case Else: Set DiKqCntwEmCnt = DiKqCnt
End Select
End Function

Function DiKqCnt(Ay, Optional Opt As EmCnt, Optional C As VbCompareMethod = vbTextCompare) As Dictionary
Dim D As Dictionary: Set D = CntzAyD(Ay, C)
Set DiKqCnt = DiKqCntwEmCnt(D, Opt)
End Function

Function DiKqCntzDrs(A As Drs, C$) As Dictionary
Set DiKqCntzDrs = DiKqCnt(ColzDrs(A, C))
End Function

