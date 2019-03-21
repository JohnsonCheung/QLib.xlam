Attribute VB_Name = "MVb_Ay_BoolAy"
Option Explicit

Enum eBoolOp
    eeOpEQ = 1
    eeOpNE = 2
    eeOpAND = 3
    eeOpOR = 4
End Enum
Enum eEqNeOp
    eeOpEQ = eBoolOp.eeOpEQ
    eeOpNE = eBoolOp.eeOpNE
End Enum
Enum eAndOrOp
    eeOpAND = eBoolOp.eeOpAND
    eeOpOR = eBoolOp.eeOpOR
End Enum

Function AndBoolAy(A() As Boolean) As Boolean
AndBoolAy = IsAllTrue(A)
End Function

Function IsAllFalse(A() As Boolean) As Boolean
Dim I
For Each I In A
    If I Then Exit Function
Next
IsAllFalse = True
End Function

Function IsAllTrue(A() As Boolean) As Boolean
Dim I
For Each I In A
    If Not I Then Exit Function
Next
IsAllTrue = True
End Function

Function IsSomTrue(A() As Boolean) As Boolean
Dim I
For Each I In A
    If I Then IsSomTrue = True: Exit Function
Next
End Function

Function OrBoolAy(A() As Boolean) As Boolean
OrBoolAy = IsSomTrue(A)
End Function


Function IsSomFalse(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If Not A(J) Then IsSomFalse = True: Exit Function
Next
End Function


Function BoolOp(BoolOpStr) As eBoolOp
Dim O As eBoolOp
Select Case UCase(BoolOpStr)
Case "AND": O = eBoolOp.eeOpAND
Case "OR": O = eBoolOp.eeOpOR
Case "EQ": O = eBoolOp.eeOpEQ
Case "NE": O = eBoolOp.eeOpNE
Case Else: Stop
End Select
BoolOp = O
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

Function IsVdtBoolOpStr(BoolOpStr$) As Boolean
IsVdtBoolOpStr = HitAy(BoolOpStr, BoolOpSy)
End Function

Function IfStr$(IfTrue As Boolean, RetStr$)
If IfTrue Then IfStr = RetStr
End Function

Property Get BoolOpSy() As String()
Static Y$(), X As Boolean
If Not X Then
    X = True
    Y = SySsl("AND OR")
End If
BoolOpSy = Y
End Property
