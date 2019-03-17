Attribute VB_Name = "Module1"
Option Explicit
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
Function FmtAsc(Optional RplNonPrtByAsc% = 8) As String()
FmtAsc = FmtSq(AscSqNoNonPrt)
End Function
