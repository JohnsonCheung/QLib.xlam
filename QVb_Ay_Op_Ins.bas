Attribute VB_Name = "QVb_Ay_Op_Ins"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Op_Ins."
Private Const Asm$ = "QVb"

Function AyInsVVAt(A, V1, V2, Optional At&)
Dim O: O = A
'AyReszCnt O, At, 2
Asg V1, O(At)
Asg V2, O(At + 1)
AyInsVVAt = O
End Function
Function AyIns(Ay)
AyIns = AyInsEle(Ay, Empty)
End Function

Function AyInsEle(Ay, ele, Optional At& = 0)
AyInsEle = AyInsAyAt(Ay, Array(ele), At)
End Function
Private Sub Z_AyInsEleAt()
Dim A(), M, At&
'--
A = Array(1, 2, 3, 4, 5)
M = "a"
At = 2
Ept = Array(1, 2, "a", 3, 4, 5)
GoSub Tst
'
Exit Sub
Tst:
'    Act = AyInsEleAt(A, M, At)
    C
    Return
End Sub
Function AyInsAy(A, B)
AyInsAy = AyInsAyAt(A, B, 0)
End Function

Function AyInsAyAt(A, B, At&)
Dim O, NB&, J&
NB = Si(B)
O = AyResz(A, At, NB)
For J = 0 To NB - 1
    Asg B(J), O(At + J)
Next
AyInsAyAt = O
End Function

Private Function AyResz(Ay, At&, Optional Cnt = 1)
Dim J&, F&, T&, U&, O, NewU&
U = UB(Ay)
NewU = U + Cnt
O = Ay
ReDim Preserve O(NewU)
Dim X&
X = NewU + At
For J& = At To U
    T = X - J
    F = T - Cnt
    Asg Ay(F), O(T)
Next
AyResz = O
End Function

Private Sub Z_AyResz()
Dim Ay(), At&, Cnt&
Ay = Array(1, 2, 3)
At = 1
Cnt = 3
Ept = Array(1, Empty, Empty, Empty, 2, 3)
Exit Sub
Tst:
'    Act = AyReszCnt(Ay, At, Cnt)
    C
    Return
End Sub

Private Sub ZZ()
Z_AyResz
MVb_AyIns:
End Sub
