Attribute VB_Name = "Module1"
Option Explicit
Const CMod$ = "Module1."

Function AscWs() As Worksheet
RgzSq AscSqNoNonPrt, NewA1(Vis:=True)
End Function

Function AscSqNoNonPrt() As Variant()
AscSqNoNonPrt = AscSqzRplNonPrt(AscSq, 8)
End Function

Function AscSq() As Variant()
Dim O(1 To 16, 1 To 16)
Dim I As Byte, J As Byte
For I = 0 To 15
For J = 0 To 15
    O(I + 1, J + 1) = Chr(I * 16 + J)
Next: Next
AscSq = O
End Function

Function AscSqzRplNonPrt(AscSq(), RplByAsc%) As Variant()
Dim O(): O = AscSq
Dim I%, J%
For I = 0 To 15
For J = 0 To 15
    If Not IsAscPrintable(Asc(O(I + 1, J + 1))) Then
        O(I + 1, J + 1) = Chr(RplByAsc)
    End If
Next: Next
AscSqzRplNonPrt = O
End Function

Sub BrwAsc()
Brw FmtAsc
End Sub
Function IsVdtAsc(AscSq) As Boolean
Select Case True
Case _
Not IsArray(AscSq), _
UBound(AscSq, 1) <> 16, _
LBound(AscSq, 1) <> 1, _
UBound(AscSq, 2) <> 16, _
LBound(AscSq, 2) <> 1
Exit Function
End Select
IsVdtAsc = True
End Function
Property Get HexDigAy() As String()
Dim J%
For J = 0 To 15: PushI HexDigAy, Hex(J): Next
End Property
Function AscAddLbl(AscSq) As Variant()
Const CSub$ = CMod & "AscAddLbl"
If Not IsVdtAsc(AscSq) Then Thw CSub, "Given AscSq is invalid.  Vdt-AscSq must 1-16 x 1-16"
Dim O(1 To 17, 1 To 17)
Dim R%, C%
For R = 2 To 17: For C = 2 To 17
    O(R, C) = AscSq(R - 1, C - 1)
Next: Next
Dim A$(): A = HexDigAy
For R = 2 To 17
    O(R, 1) = A(R - 2)
Next
For C = 2 To 17
    O(1, C) = A(C - 2)
Next
O(1, 1) = " "
AscAddLbl = O
End Function
Function FmtAsc(Optional RplNonPrtByAsc% = 8) As String()
FmtAsc = FmtSq(AscAddLbl(AscSqNoNonPrt))
End Function

