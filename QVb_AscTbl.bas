Attribute VB_Name = "QVb_AscTbl"
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_AscTbl."
Const Asc99% = &H99
Function Chr99$()
Chr99 = Chr(Asc99)
End Function
Function AscWs() As Worksheet
RgzSq AscSqOfNoNonPrt, NewA1(Vis:=True)
End Function

Property Get AscSqOfNoNonPrt() As Variant()
AscSqOfNoNonPrt = AscSqRplNonPrt(AscSq, 8)
End Property

Property Get AscSq() As Variant()
Dim O(1 To 16, 1 To 16)
Dim I As Byte, J As Byte
For I = 0 To 15
For J = 0 To 15
    O(I + 1, J + 1) = Chr(I * 16 + J)
Next: Next
AscSq = O
End Property

Function AscSqRplNonPrt(AscSq(), RplByAsc%) As Variant()
Dim O(): O = AscSq
Dim I%, J%
For I = 0 To 15
For J = 0 To 15
    If Not IsAscPrintable(Asc(O(I + 1, J + 1))) Then
        O(I + 1, J + 1) = Chr(RplByAsc)
    End If
Next: Next
AscSqRplNonPrt = O
End Function

Sub BrwAsc()
Brw FmtAsc
End Sub
Function IsVdtAsc(AscSq()) As Boolean
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

Function AscSqAddLbl(AscSq()) As Variant()
Const CSub$ = CMod & "AscSqAddLbl"
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
AscSqAddLbl = O
End Function

Function FmtAsc(Optional RplNonPrtByAsc% = 8) As String()
FmtAsc = FmtSq(AscSqAddLbl(AscSqOfNoNonPrt))
End Function

Function FmtSq(Sq(), Optional SepChr$ = " ") As String()
If Si(Sq) = 0 Then Exit Function
With RRCCzSq(Sq)
Dim I&
For I = .R1 To .R2
    PushI FmtSq, Jn(DrzSqr(Sq(), I), SepChr)
Next
End With
End Function

Property Get FmtAscSq() As String()
FmtAscSq = FmtSq(AscSqOfNoNonPrt)
End Property

Sub DmpAsc(S, Optional MaxLen& = 100)
Dim J&, C$
Debug.Print "Len=" & Len(S)
For J = 1 To Min(MaxLen, Len(S))
    C = Mid(S, J, 1)
    Debug.Print J, Asc(C), C
Next
End Sub

Sub DmpAscSq()
Dmp FmtAscSq
End Sub


Function RRCCzSq(Sq()) As RRCC
Set RRCCzSq = New RRCC
With RRCCzSq
    .R1 = LBound(Sq(), 1)
    .R2 = UBound(Sq(), 1)
    .C1 = LBound(Sq(), 2)
    .C2 = UBound(Sq(), 2)
End With
End Function


