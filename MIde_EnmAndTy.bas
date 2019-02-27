Attribute VB_Name = "MIde_EnmAndTy"
Option Explicit
Type Pos
    Cno1 As Long
    Cno2 As Long
End Type
Type LinPos
    Lno As Long
    Pos As Pos
End Type
Type MdPos
    Md As CodeModule
    Pos As LinPos
End Type
Function LinPosLno(Lno&) As LinPos
LinPosLno.Lno = Lno
End Function
Function Pos(C1, C2) As Pos
If C1 > 0 Then Pos.Cno1 = C1
If C1 > 0 And C2 >= C1 Then Pos.Cno2 = C2
End Function
Function MdPos(Md As CodeModule, Pos As LinPos) As MdPos
With MdPos
    Set .Md = Md
    .Pos = Pos
End With
End Function

Function LinPos(Lno, P As Pos) As LinPos
With LinPos
    .Lno = Lno
    .Pos = P
End With
End Function
