Attribute VB_Name = "QIde_F_Pos"
Option Explicit
Option Compare Text
Type Pos
    Cno1 As Long
    Cno2 As Long
End Type
Type MdPos
    Md As CodeModule
    RRCC As RRCC
End Type
Type MdPoses: N As Long: Ay() As MdPos: End Type
Function Pos(C1, C2) As Pos
If C1 <= 0 Then Exit Function
If C2 <= 0 Then Exit Function
If C1 > C2 Then Exit Function
Pos.Cno2 = C2
Pos.Cno1 = C1
End Function
Function PoszSS(S, SubStr) As Pos
Dim P%: P = InStr(S, SubStr)
If P > 0 Then PoszSS = Pos(P, P + Len(SubStr))
End Function
Function EmpPos() As Pos
End Function
Function SubStrPos(S, SubStr) As Pos

End Function

Function PoszSubStr(S, SubStr) As Pos
Dim P&: P = InStr(S, SubStr): If P = 0 Then Exit Function
PoszSubStr = Pos(P, P + Len(SubStr) - 1)
End Function


'
