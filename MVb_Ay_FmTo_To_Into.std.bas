Attribute VB_Name = "MVb_Ay_FmTo_To_Into"
Option Explicit


Function IntozAy(Into, Ay)
If TypeName(Ay) = TypeName(Into) Then
    IntozAy = Ay
    Exit Function
End If
IntozAy = AyCln(Into)
Dim I
For Each I In Itr(Ay)
    Push IntozAy, I
Next
End Function



Function IntozItrNy(Into, Itr, Ny$())
Dim O: O = Into: Erase O
Dim Obj
For Each Obj In Itr
    If HasEle(Ny, ObjNm(Obj)) Then
        PushObj O, Obj
    End If
Next
IntozItrNy = O
End Function
Function IntozItr(Into, Itr)
Dim O: O = Into: Erase O
Dim Obj
For Each Obj In Itr
    Push O, Obj
Next
IntozItr = O
End Function

