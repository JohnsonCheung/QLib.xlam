Attribute VB_Name = "QVb_Ay_BoolAy"
Option Explicit

Private Const CMod$ = "MVb_Ay_BoolAy."
Private Const Asm$ = "QVb"
Enum EmBoolOp
    EiOpEQ = 1
    EiOpNE = 2
    EiOpAND = 3
    EiOpOR = 4
End Enum
Enum EmEqNeOp
    EiOpEQ = EmBoolOp.EiOpEQ
    EiOpNE = EmBoolOp.EiOpNE
End Enum
Enum EmAndOrOp
    EiOpAND = EmBoolOp.EiOpAND
    EiOpOR = EmBoolOp.EiOpOR
End Enum

Function AndBoolAy(A() As Boolean) As Boolean
AndBoolAy = IsAllTruezB(A)
End Function

Function BoolOp(BoolOpStr$) As EmBoolOp
Dim O As EmBoolOp
Select Case UCase(BoolOpStr)
Case "AND": O = EmBoolOp.EiOpAND
Case "OR": O = EmBoolOp.EiOpOR
Case "EQ": O = EmBoolOp.EiOpEQ
Case "NE": O = EmBoolOp.EiOpNE
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
OrBoolAy = IsSomTrue(A)
End Function

Property Get BoolOpSy() As String()
Static Y$(), X As Boolean
If Not X Then
    X = True
    Y = SyzSsLin("AND OR")
End If
BoolOpSy = Y
End Property
