Attribute VB_Name = "MVb_Str_Brk"
Option Explicit
Const CMod$ = "MVb_Str_Brk."
Sub AsgBrk1Dot(S, OA, OB, Optional NoTrim As Boolean)
AsgS1S2 Brk1Dot(S), OA, OB
End Sub
Sub AsgBrkDot(S, OA, OB, Optional NoTrim As Boolean)
AsgS1S2 BrkDot(S), OA, OB
End Sub
Function Brk1Dot(S, Optional NoTrim As Boolean) As S1S2
Set Brk1Dot = Brk1(S, ".", NoTrim)
End Function
Function Brk2Dot(S, Optional NoTrim As Boolean) As S1S2
Set Brk2Dot = Brk2(S, ".", NoTrim)
End Function
Function BrkDot(S, Optional NoTrim As Boolean) As S1S2
Set BrkDot = Brk(S, ".", NoTrim)
End Function
Function Brk(A, Sep, Optional NoTrim As Boolean) As S1S2
Const CSub$ = CMod & "Brk"
Dim P&: P = InStr(A, Sep)
If P = 0 Then Thw CSub, "{S} does not contains {Sep}", "S Sep", A, Sep
Set Brk = BrkAtSep(A, P, Sep, NoTrim)
End Function

Function Brk1(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
If P = 0 Then Set Brk1 = S1S2(A, "", NoTrim): Exit Function
Set Brk1 = Brk1At(A, P, Sep, NoTrim)
End Function

Sub AsgBrk1(A, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
AsgS1S2 Brk1(A, Sep, NoTrim), O1, O2
End Sub

Function Brk1Rev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
If P = 0 Then Set Brk1Rev = S1S2(A, "", NoTrim): Exit Function
Set Brk1Rev = Brk1At(A, P, Sep, NoTrim)
End Function

Function Brk2(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Set Brk2 = Brk2__(A, P, Sep, NoTrim)
End Function

Function Brk2__(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    If NoTrim Then
        Set Brk2__ = S1S2("", A)
    Else
        Set Brk2__ = S1S2("", Trim(A))
    End If
    Exit Function
End If
Set Brk2__ = Brk1At(A, P, Sep, NoTrim)
End Function

Sub AsgBrk2(A, Sep$, O1, O2, Optional NoTrim As Boolean)
AsgS1S2 Brk2(A, Sep, NoTrim), O1, O2
End Sub

Function Brk2Rev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
Set Brk2Rev = Brk2__(A, P, Sep, NoTrim)
End Function

Sub AsgBrk(A, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
AsgBrkAt A, InStr(A, Sep), Sep, O1, O2, NoTrim
End Sub

Private Function BrkAtSep(A, P&, Sep, NoTrim As Boolean) As S1S2
Dim S1$, S2$
S1 = Left(A, P - 1)
S2 = Mid(A, P + Len(Sep))
Set BrkAtSep = S1S2(S1, S2, NoTrim)
End Function

Function Brk1At(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    Set Brk1At = S1S2(A, "", NoTrim)
Else
    Set Brk1At = BrkAtSep(A, P, Sep, NoTrim)
End If
End Function

Sub AsgBrkAt(A, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
Const CSub$ = CMod & "AsgBrkAt"
If At = 0 Then
    Thw CSub, "String does not have Sep", "Str Sep At NoTrim", A, Sep, At, NoTrim
    Exit Sub
End If
O1 = Left(A, At - 1)
O2 = Mid(A, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub

Function BrkBoth(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
If P = 0 Then
    Set BrkBoth = S1S2(A, A, NoTrim)
    Exit Function
End If
Set BrkBoth = Brk1At(A, P, Sep, NoTrim)
End Function

Sub AsgBrkQuote(QuoteStr, O1$, O2$)
AsgS1S2 BrkQuote(QuoteStr), O1, O2
End Sub

Function BrkRev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
If P = 0 Then Err.Raise "BrkRev: Str[" & A & "] does not contains Sep[" & Sep & "]"
BrkRev = Brk1At(A, P, Len(Sep), NoTrim)
End Function

Sub AsgBrk1At(A, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
If At = 0 Then
    O1 = A
    O2 = ""
    Exit Sub
End If
O1 = Left(A, At - 1)
O2 = Mid(A, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub

Private Sub ZZ_Brk1Rev()
Dim S1$, S2$, ExpS1$, ExpS2$, A$
A = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(A, "---")
    S1 = .S1
    S2 = .S2
End With
Ass S1 = ExpS1
Ass S2 = ExpS2
End Sub

Private Sub Z_Brk1Rev()
Dim S1$, S2$, ExpS1$, ExpS2$, A$
A = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(A, "---")
    S1 = .S1
    S2 = .S2
End With
Ass S1 = ExpS1
Ass S2 = ExpS2
End Sub

Private Sub Z()
Z_Brk1Rev
MVb_Str_Brk:
End Sub
