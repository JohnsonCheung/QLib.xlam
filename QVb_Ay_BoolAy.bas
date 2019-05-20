Attribute VB_Name = "QVb_Ay_BoolAy"
Option Compare Text
Option Explicit

Private Const CMod$ = "MVb_Ay_BoolAy."
Private Const Asm$ = "QVb"
Enum EmBoolOp
    EiOpEq = 1
    EiOpNe = 2
    EiOpAnd = 3
    EiOpOr = 4
End Enum
Enum EmEqNeOp
    EiOpEq = EmBoolOp.EiOpEq
    EiOpNe = EmBoolOp.EiOpNe
End Enum
Enum EmAndOrOp
    EiOpAnd = EmBoolOp.EiOpAnd
    EiOpOr = EmBoolOp.EiOpOr
End Enum

Function AndBoolAy(A() As Boolean) As Boolean
AndBoolAy = IsAllTruezB(A)
End Function

Function BoolOp(BoolOpStr$) As EmBoolOp
Dim O As EmBoolOp
Select Case UCase(BoolOpStr)
Case "AND": O = EmBoolOp.EiOpAnd
Case "OR": O = EmBoolOp.EiOpOr
Case "EQ": O = EmBoolOp.EiOpEq
Case "NE": O = EmBoolOp.EiOpNe
Case Else: Stop
End Select
BoolOp = O
End Function

Function IfStr$(IfTrue As Boolean, RetStr$)
If IfTrue Then IfStr = RetStr
End Function

Function IsAllFalsezB(A() As Boolean) As Boolean
Dim I
For Each I In A
    If I Then Exit Function
Next
IsAllFalsezB = True
End Function

Function IsAllTruezB(A() As Boolean) As Boolean
Dim I
For Each I In A
    If Not I Then Exit Function
Next
IsAllTruezB = True
End Function

Function IsAndOrStr(A$) As Boolean
Select Case UCase(A)
Case "AND", "OR": IsAndOrStr = True
End Select
End Function

Function IsEqNeStr(A$) As Boolean
Select Case UCase(A)
Case "EQ", "NE": IsEqNeStr = True
End Select
End Function

Function IsSomFalsezB(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If Not A(J) Then IsSomFalsezB = True: Exit Function
Next
End Function

Function IsSomTruezB(A() As Boolean) As Boolean
Dim I
For Each I In A
    If I Then IsSomTruezB = True: Exit Function
Next
End Function

Function IsVdtBoolOpStr(BoolOpStr$) As Boolean
IsVdtBoolOpStr = HitAy(BoolOpStr, BoolOpSy)
End Function

Function OrBoolAy(A() As Boolean) As Boolean
OrBoolAy = IsSomTruezB(A)
End Function

Property Get BoolOpSy() As String()
Static Y$(), X As Boolean
If Not X Then
    X = True
    Y = SyzSS("AND OR")
End If
BoolOpSy = Y
End Property
