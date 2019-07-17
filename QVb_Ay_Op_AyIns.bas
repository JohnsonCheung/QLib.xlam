Attribute VB_Name = "QVb_Ay_Op_AyIns"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Op_Ins."
Private Const Asm$ = "QVb"

Function Ins2Ele(Ay, E1, E2, Optional At&)
Ins2Ele = InsAy(Ay, Array(E1, E2), At)
End Function

Function InsEle(Ay, Optional Ele, Optional At& = 0)
InsEle = InsAy(Ay, Array(Ele), At)
End Function

Private Sub Z_InsEle()
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
    Act = InsEle(A, M, At)
    C
    Return
End Sub

Function InsAy(AyA, AyB, At&)
Dim O, NB&, J&
NB = Si(AyB)
O = ResiAt(AyA, At, NB)
For J = 0 To NB - 1
    Asg AyB(J), O(At + J)
Next
InsAy = O
End Function

Private Sub Z_ReBaseR()
Dim Ay&(): ReDim Ay(9)
Dim J%: For J = 0 To 9: Ay(J) = J: Next
Dim Act: Act = ReBase(Ay, 100)
Debug.Assert LBound(Act) = 100
Debug.Assert UBound(Act) = 109
For J = 100 To 109
    Debug.Assert Act(J) = J - 100
Next
End Sub
Function ReBase(Ay, FmIx)
'Fm Ay : assume Ay-LBound is 0
'Ret   : new re-based ay of LBound is @FmIx preserve.  Note standard redim preserve X(F To T) does not work  @@
Dim ToIx&: ToIx = FmIx + UB(Ay)
Dim O:        O = Ay: ReDim O(FmIx To ToIx)
Dim J&:       J = FmIx
Dim V:            For Each V In Ay
    Asg V, O(J)
J = J + 1
Next
ReBase = O
End Function

Function ResiAt(Ay, At&, Optional Cnt = 1)
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
ResiAt = O
End Function

Private Sub Z_Resi()
Dim Ay(), At&, Cnt&
Ay = Array(1, 2, 3)
At = 1
Cnt = 3
Ept = Array(1, Empty, Empty, Empty, 2, 3)
Exit Sub
Tst:
'    Act = ResiCnt(Ay, At, Cnt)
    C
    Return
End Sub

Private Sub Z()
Z_Resi
MVb_AyIns:
End Sub
