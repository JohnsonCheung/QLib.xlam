Attribute VB_Name = "MVb_Ay_Op_Cnt"
Option Explicit

Function CntDryWhGt1zAy(A) As Variant()
CntDryWhGt1zAy = CntDryWhGt1(DryzDic(CntDic(A)))
End Function

Function CntDryzAy(A) As Variant()
CntDryzAy = DryzDic(CntDic(A))
End Function

Private Sub Z_CntDryzAy()
Dim A$()
A = SplitSpc("a a a b c b")
Ept = Array(Array("a", 3), Array("b", 2), Array("c", 1))
GoSub Tst
Exit Sub
Tst:
    Act = CntDryzAy(A)
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
Debug.Print CntSiLin(SrczPj(CurPj))
End Sub
Function CntSiLin$(Ay)
CntSiLin = "AyCntSi(" & Si(Ay) & "." & SumSi(Ay) & ")"
End Function
Private Sub Z()
Z_CntDryzAy
MVb_AyCnt:
End Sub
