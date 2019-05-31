Attribute VB_Name = "QVb_Ay_Op_IxAy"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Op_Ixy."
Private Const Asm$ = "QVb"
Type NumPos
    Num As Long
    Pos As Long
End Type

Private Sub ZZ_AsgIxy()
Dim Fny$(), FF$, A%, B%, C%, EA%, EB%, EC%
GoSub T1
Exit Sub
T1:
    Fny = SyzSS("A B C")
    FF = "C B A"
    EA = 0
    EB = 1
    EC = 2
    GoTo Tst
Tst:
    AsgIxy Fny, FF, "C,B,A"
    Debug.Print A = EA
    Debug.Print B = EB
    Debug.Print C = EC
    Return
End Sub
Sub AsgIxy(Fny$(), FF$, ParamArray OAp())
Dim F, J%, I&
For Each F In SyzSS(FF)
    I = IxzAy(Fny, F): If I < 0 Then Thw CSub, "F in FF not found in Fny", "Fny FF, F", Fny, FF, F
    OAp(J) = IxzAy(Fny, F)
    J = J + 1
Next
End Sub

Function IxzAy&(Ay, Itm, Optional FmIx& = 0)
Dim J&
For J = FmIx To UB(Ay)
    If Ay(J) = Itm Then IxzAy = J: Exit Function
Next
IxzAy = -1
End Function
Function IxyzU(U&) As Long()
Dim O&(), J&
ReDim O(U)
For J = 0 To U
    O(J) = J
Next
IxyzU = O
End Function

Function GrpEndIxyOfEmp(Ay) As Long()
Dim J&, Fst As Boolean, U&
U = UB(Ay)
For J = 1 To U
    If Not IsEmpty(Ay(J)) Then
        PushI GrpEndIxyOfEmp, J - 1
    End If
Next
PushI GrpEndIxyOfEmp, U
End Function

Function IxyzEle(Ay, Ele) As Long()
Dim J&, V
For Each V In Itr(Ay)
    If V = Ele Then PushI IxyzEle, J
    J = J + 1
Next
End Function
Function Ixy(Ay, SubAy, Optional ThwNotFnd As Boolean) As Long()
Dim I, HasNegIx As Boolean, Ix&
For Each I In Itr(SubAy)
    Ix = IxzAy(Ay, I)
    If Ix = -1 Then HasNegIx = True
    PushI Ixy, Ix
Next
If ThwNotFnd Then
    If HasNegIx Then
        Thw CSub, "There is negative index", "Ay SubAy Ixy", Ay, SubAy, Ixy
    End If
End If
End Function

Sub AsgItmAyIxay(Ay, Ixy&(), ParamArray OItmAp())
Dim J&
For J = 0 To UB(Ixy)
    Asg Ay(Ixy(J)), OItmAp(J)
Next
End Sub

Function IntIxy(Ay, SubAy) As Integer()
Dim J&, U&
For J = 0 To UB(SubAy)
    PushI IntIxy, IxzAy(Ay, SubAy(J))
Next
End Function
Function IxyzDup(AyWithDup) As Long()
Dim A As Aset: Set A = AsetzAy(AywDup(AyWithDup))
If A.IsEmp Then Exit Function
Dim J&
For J = 0 To UB(AyWithDup)
    If A.Has(AyWithDup(J)) Then PushI IxyzDup, J
Next
End Function


