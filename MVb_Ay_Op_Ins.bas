Attribute VB_Name = "MVb_Ay_Op_Ins"
Option Explicit

Function AyInsVVAt(A, V1, V2, Optional At&)
Dim O: O = A
'AyRgzReszCnt O, At, 2
Asg V1, O(At)
Asg V2, O(At + 1)
AyInsVVAt = O
End Function
Function AyIns(Ay)
AyIns = AyInsItm(Ay, Empty)
End Function

Function AyInsItm(Ay, Itm, Optional At& = 0)
AyInsItm = AyInsAyAt(Ay, Array(Itm), At)
End Function
Private Sub Z_AyInsItmAt()
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
'    Act = AyInsItmAt(A, M, At)
    C
    Return
End Sub
Function AyInsAy(A, B)
AyInsAy = AyInsAyAt(A, B, 0)
End Function

Function AyInsAyAt(A, B, At&)
Dim O, NB&, J&
NB = Sz(B)
O = AyRgzResz(A, At, NB)
For J = 0 To NB - 1
    Asg B(J), O(At + J)
Next
AyInsAyAt = O
End Function

Private Function AyRgzResz(Ay, At&, Optional Cnt = 1)
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
AyRgzResz = O
End Function

Private Sub Z_AyRgzResz()
Dim Ay(), At&, Cnt&
Ay = Array(1, 2, 3)
At = 1
Cnt = 3
Ept = Array(1, Empty, Empty, Empty, 2, 3)
Exit Sub
Tst:
'    Act = AyRgzReszCnt(Ay, At, Cnt)
    C
    Return
End Sub

Private Sub Z()
Z_AyRgzResz
MVb_AyIns:
End Sub
