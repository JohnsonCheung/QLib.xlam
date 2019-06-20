Attribute VB_Name = "QVb_Ay_FmTo_To_Into"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_FmTo_To_Into."
Private Const Asm$ = "QVb"

Function IntozAy(Into, Ay)
If IsEqTy(Ay, Into) Then
    IntozAy = Ay
    Exit Function
End If
IntozAy = ResiU(Into)
Dim I
For Each I In Itr(Ay)
    Push IntozAy, I
Next
End Function

Function IntozItrNy(Into$, Itr, Ny$())
Dim O: O = Into: Erase O
Dim Obj
For Each Obj In Itr
    If HasEle(Ny, ObjNm(Obj)) Then
        PushObj O, Obj
    End If
Next
IntozItrNy = O
End Function
Function SyzItr(Itr) As String()
SyzItr = IntozItr(EmpSy, Itr)
End Function
Function IntozItr(Into, Itr)
Dim O: O = Into: Erase O
Dim Obj
For Each Obj In Itr
    Push O, Obj
Next
IntozItr = O
End Function

