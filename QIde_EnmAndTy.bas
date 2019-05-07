Attribute VB_Name = "QIde_EnmAndTy"
Option Explicit
Private Const CMod$ = "MIde_EnmAndTy."
Private Const Asm$ = "QIde"
Function CvLinPos(A) As LinPos
Set CvLinPos = A
End Function
Function LinPos(Lno, Optional Cno1 = 0, Optional Cno2 = 0) As LinPos
Dim O As New LinPos
Set LinPos = O.Init(Lno, Pos(Cno1, Cno2))
End Function
Function SubStrPos(S, SubStr) As Pos
Dim P&: P = InStr(S, SubStr): If P = 0 Then Set SubStrPos = New Pos: Exit Function
Set SubStrPos = Pos(P, P + Len(SubStr) - 1)
End Function
Function Pos(Optional Cno1 = 0, Optional Cno2 = 0) As Pos
Dim O As New Pos
Set Pos = O.Init(Cno1, Cno2)
End Function
Function MdPosStr$(A As MdPos)
With A
    Dim B$
    With .Pos.Pos
        If .Cno1 > 0 Then B = " " & .Cno1 & " " & .Cno2
    End With
    MdPosStr = "MdPos " & MdNm(A.Md) & A.Pos.Lno & B
End With
End Function

Function MdPos(Md As CodeModule, Pos As LinPos) As MdPos
Dim O As New MdPos
Set MdPos = O.Init(Md, Pos)
End Function

