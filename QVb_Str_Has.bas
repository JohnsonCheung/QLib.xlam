Attribute VB_Name = "QVb_Str_Has"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_S_Has."
Private Const Asm$ = "QVb"
Enum EmCas
    EiIgn = 0
    EiSen = 1
End Enum

Function HasDot(S) As Boolean
HasDot = HasSubStr(S, ".")
End Function
Function HasSngQ(S) As Boolean
HasSngQ = InStr(S, vbSngQ)
End Function
Function HasDblQ(S) As Boolean
HasDblQ = InStr(S, vbDblQ)
End Function

Function RmvBetDblQ$(S)
Dim P&: P = InStr(S, vbDblQ)
Dim O$: O = S
While P > 0
    Dim J%: J = J + 1: If J > 10000 Then Stop
    Dim P1&: P1 = InStr(P + 1, O, vbDblQ): If P1 = 0 Then Stop
    O = Left(O, P - 1) & Mid(O, P1 + 1)
    P = InStr(P + 1, O, vbDblQ)
Wend
RmvBetDblQ = O
End Function
Function HasSngDblQ(S) As Boolean
If HasSngQ(S) Then
    If HasDblQ(S) Then
        HasSngDblQ = True
    End If
End If
End Function
Function HasSubStr(S, SubStr, Optional Cpr As VbCompareMethod) As Boolean
HasSubStr = InStr(1, S, SubStr, Cpr) > 0
End Function

Function HasCrLf(S) As Boolean
HasCrLf = HasSubStr(S, vbCrLf)
End Function

Function HasHyphen(S) As Boolean
HasHyphen = HasSubStr(S, "-")
End Function

Function HasPound(S) As Boolean
HasPound = InStr(S, "#") > 0
End Function

Function HasSpc(S) As Boolean
HasSpc = InStr(S, " ") > 0
End Function

Function HasSqBkt(S) As Boolean
HasSqBkt = FstChr(S) = "[" And LasChr(S) = "]"
End Function

Function HasChrList(S, ChrList$, Optional Cpr As VbCompareMethod) As Boolean
Dim J%
For J = 1 To Len(ChrList)
    If HasSubStr(S, Mid(ChrList, J, 1), Cpr) Then HasChrList = True: Exit Function
Next
End Function

Function HasSubStrAy(S, SubStrAy$()) As Boolean
Dim SubStr
For Each SubStr In SubStrAy
    If HasSubStr(S, SubStr) Then HasSubStrAy = True: Exit Function
Next
End Function
Function HasTT(S, T1, T2) As Boolean
HasTT = Has2T(S, T1, T2)
End Function

Function HasVbar(S) As Boolean
HasVbar = HasSubStr(S, "|")
End Function
