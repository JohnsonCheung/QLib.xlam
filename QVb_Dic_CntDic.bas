Attribute VB_Name = "QVb_Dic_CntDic"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Dic_CntDic."
Private Const Asm$ = "QVb"
Function FmtCntDic(Ay, Optional Opt As EmCnt) As String()
FmtCntDic = FmtS12s(SwapS12s(S12szDic(CntDic(Ay, Opt))), N1:="Cnt", N2:="Mth")
End Function

Function CntzAyD(Ay, Optional C As VbCompareMethod = vbTextCompare) As Dictionary
'Ret : :CntDic #Cnt-Ay-Dic ! Cnt-Ay-ret-as-cntDic.  :CntDic is val is a cnt (num).  %Cnt always :CntDic
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
Function CntDicwDup(CntDic As Dictionary) As Dictionary
Set CntDicwDup = New Dictionary
Dim Cnt&, K
For Each K In CntDic.Keys
    Cnt = CntDic(K)
    If Cnt > 1 Then CntDicwDup.Add K, Cnt
Next
End Function

Function CntDicwSng(CntDic As Dictionary) As Dictionary
Set CntDicwSng = New Dictionary
Dim Cnt&, K
For Each K In CntDic.Keys
    Cnt = CntDic(K)
    If Cnt = 1 Then CntDicwSng.Add K, Cnt
Next
End Function
Function CntDicwEmCnt(CntDic As Dictionary, B As EmCnt) As Dictionary
Select Case B
Case EmCnt.EiCntDup: Set CntDicwEmCnt = CntDicwDup(CntDic)
Case EmCnt.EiCntSng: Set CntDicwEmCnt = CntDicwSng(CntDic)
Case Else: Set CntDicwEmCnt = CntDic
End Select
End Function

Function CntDic(Ay, Optional Opt As EmCnt, Optional C As VbCompareMethod = vbTextCompare) As Dictionary
Dim D As Dictionary: Set D = CntzAyD(Ay, C)
Set CntDic = CntDicwEmCnt(D, Opt)
End Function

Function CntDiczDrs(A As Drs, C$) As Dictionary
Set CntDiczDrs = CntDic(ColzDrs(A, C))
End Function

