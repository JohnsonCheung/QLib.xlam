Attribute VB_Name = "MxAyCnt"
Option Compare Text
Option Explicit
Const CLib$ = "QVb."
Const CMod$ = CLib & "MxAyCnt."

Function CntDyWhGt1zAy(A) As Variant()
CntDyWhGt1zAy = CntDyWhGt1(DyoDic(DiKqCnt(A)))
End Function

Function CntDyoAy(A) As Variant()
CntDyoAy = DyoDic(DiKqCnt(A))
End Function

Private Sub Z_CntDyoAy()
Dim A$()
A = SplitSpc("a a a b c b")
Ept = Array(Array("a", 3), Array("b", 2), Array("c", 1))
GoSub Tst
Exit Sub
Tst:
    Act = CntDyoAy(A)
    Ass IsEqAy(Act, Ept)
    Return
End Sub

Function SumSi&(Ay)
Dim I, O&
For Each I In Itr(Ay)
    O = O + Len(I)
Next
SumSi = O
End Function
Private Sub Z_CntSiLin()
Debug.Print CntSiLin(SrczP(CPj))
End Sub
Function CntSiLin(Ay)
CntSiLin = "AyCntSi(" & Si(Ay) & "." & SumSi(Ay) & ")"
End Function
Private Sub Z()
Z_CntDyoAy
MVb_AyCnt:
End Sub