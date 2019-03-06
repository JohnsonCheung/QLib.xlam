Attribute VB_Name = "MVb_Ay_Op_IxAy"
Option Explicit
Type NumPos
    Num As Long
    Pos As Long
End Type

Function IxzAy&(Ay, Itm, Optional FmIx& = 0)
Dim J&
For J = FmIx To UB(Ay)
    If Ay(J) = Itm Then IxzAy = J: Exit Function
Next
IxzAy = -1
End Function
Function IxAyU(U&) As Long()
Dim O&(), J&
ReDim O(U)
For J = 0 To U
    O(J) = J
Next
IxAyU = O
End Function

Function IxAy(Ay, SubAy, Optional ThwNotFnd As Boolean) As Long()
Dim I, HasNegIx As Boolean, Ix&
For Each I In Itr(SubAy)
    Ix = IxzAy(Ay, I)
    If Ix = -1 Then HasNegIx = True
    PushI IxAy, Ix
Next
If ThwNotFnd Then
    If HasNegIx Then
        Thw CSub, "There is negative index", "Ay SubAy IxAy", Ay, SubAy, IxAy
    End If
End If
End Function

Sub AsgItmAyIxay(Ay, IxAy&(), ParamArray OItmAp())
Dim J&
For J = 0 To UB(IxAy)
    Asg Ay(IxAy(J)), OItmAp(J)
Next
End Sub

Function IntIxAy(Ay, SubAy) As Integer()
Dim J&, U&
For J = 0 To UB(SubAy)
    PushI IntIxAy, IxzAy(Ay, SubAy(J))
Next
End Function
Function IxAyzDup(AyWithDup) As Long()
Dim A As Aset: Set A = AsetzAy(AywDup(AyWithDup))
If A.IsEmp Then Exit Function
Dim J&
For J = 0 To UB(AyWithDup)
    If A.Has(AyWithDup(J)) Then PushI IxAyzDup, J
Next
End Function

