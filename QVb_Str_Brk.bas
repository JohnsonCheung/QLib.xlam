Attribute VB_Name = "QVb_Str_Brk"
Option Compare Text
Option Explicit
Private Const Asm$ = "QVb"
Private Const CMod$ = "MVb_Str_Brk."
Sub AsgBrkSpc(S, OA$, OB$, Optional NoTrim As Boolean)
AsgS1S2 BrkSpc(S), OA, OB
End Sub
Sub AsgBrk1Dot(S, OA$, OB$, Optional NoTrim As Boolean)
AsgS1S2 Brk1Dot(S), OA, OB
End Sub
Sub AsgBrkDot(S, OA$, OB$, Optional NoTrim As Boolean)
AsgS1S2 BrkDot(S), OA, OB
End Sub
Function Brk1Dot(S, Optional NoTrim As Boolean) As S1S2
Brk1Dot = Brk1(S, ".", NoTrim)
End Function
Function Brk2Dot(S, Optional NoTrim As Boolean) As S1S2
Brk2Dot = Brk2(S, ".", NoTrim)
End Function
Function BrkDot(S, Optional NoTrim As Boolean) As S1S2
BrkDot = Brk(S, ".", NoTrim)
End Function
Function BrkSpc(S) As S1S2
BrkSpc = Brk(S, " ")
End Function
Function Brk(S, Sep$, Optional NoTrim As Boolean) As S1S2
Const CSub$ = CMod & "Brk"
Dim P&: P = InStr(S, Sep)
If P = 0 Then Thw CSub, "{S} does not contains {Sep}", "S Sep", S, Sep
Brk = BrkAtSep(S, P, Sep, NoTrim)
End Function

Function Brk1(S, Sep$, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then Brk1 = S1S2(S, "", NoTrim): Exit Function
Brk1 = Brk1At(S, P, Sep, NoTrim)
End Function

Sub AsgBrk1(S, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
AsgS1S2 Brk1(S, Sep, NoTrim), O1, O2
End Sub

Function Brk1Rev(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(S, Sep)
If P = 0 Then Brk1Rev = S1S2(S, "", NoTrim): Exit Function
Brk1Rev = Brk1At(S, P, Sep, NoTrim)
End Function

Function Brk2(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
Brk2 = Brk2__(S, P, Sep, NoTrim)
End Function

Function Brk2__(S, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    If NoTrim Then
        Brk2__ = S1S2("", S)
    Else
        Brk2__ = S1S2("", Trim(S))
    End If
    Exit Function
End If
Brk2__ = Brk1At(S, P, Sep, NoTrim)
End Function

Sub AsgBrk2(S, Sep$, O1, O2, Optional NoTrim As Boolean)
AsgS1S2 Brk2(S, Sep, NoTrim), O1, O2
End Sub

Function Brk2Rev(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(S, Sep)
Brk2Rev = Brk2__(S, P, Sep, NoTrim)
End Function

Sub AsgBrk(S, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
AsgBrkAt S, InStr(S, Sep), Sep, O1, O2, NoTrim
End Sub

Private Function BrkAtSep(S, P&, Sep, NoTrim As Boolean) As S1S2
Dim S1$, S2$
S1 = Left(S, P - 1)
S2 = Mid(S, P + Len(Sep))
BrkAtSep = S1S2(S1, S2, NoTrim)
End Function

Function Brk1At(S, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    Brk1At = S1S2(S, "", NoTrim)
Else
    Brk1At = BrkAtSep(S, P, Sep, NoTrim)
End If
End Function

Sub AsgBrkAt(S, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
Const CSub$ = CMod & "AsgBrkAt"
If At = 0 Then
    Thw CSub, "String does not have Sep", "Str Sep At NoTrim", S, Sep, At, NoTrim
    Exit Sub
End If
O1 = Left(S, At - 1)
O2 = Mid(S, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub

Function BrkBoth(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(S, Sep)
If P = 0 Then
    BrkBoth = S1S2(S, S, NoTrim)
    Exit Function
End If
BrkBoth = Brk1At(S, P, Sep, NoTrim)
End Function

Sub AsgBrkQuote(QuoteStr$, O1$, O2$)
AsgS1S2 BrkQuote(QuoteStr), O1, O2
End Sub

Function BrkRev(S, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(S, Sep)
If P = 0 Then Err.Raise "BrkRev: Str[" & S & "] does not contains Sep[" & Sep & "]"
BrkRev = Brk1At(S, P, Len(Sep), NoTrim)
End Function

Sub AsgBrk1At(S, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
If At = 0 Then
    O1 = S
    O2 = ""
    Exit Sub
End If
O1 = Left(S, At - 1)
O2 = Mid(S, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub

Private Sub ZZ_Brk1Rev()
Dim S1$, S2$, ExpS1$, ExpS2$, S
S = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(S, "---")
    S1 = .S1
    S2 = .S2
End With
Ass S1 = ExpS1
Ass S2 = ExpS2
End Sub

Private Sub Z_Brk1Rev()
Dim S1$, S2$, ExpS1$, ExpS2$, S
S = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(S, "---")
    S1 = .S1
    S2 = .S2
End With
Ass S1 = ExpS1
Ass S2 = ExpS2
End Sub

Private Sub ZZ()
Z_Brk1Rev
MVb_Str_Brk:
End Sub
