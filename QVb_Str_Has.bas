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

Function HasSubStr(S, SubStr, Optional Cmpr As VbCompareMethod) As Boolean
HasSubStr = InStr(1, S, SubStr, Cmpr) > 0
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

Function HasChrList(S, ChrList$, Optional Cmpr As VbCompareMethod) As Boolean
Dim J%
For J = 1 To Len(ChrList)
    If HasSubStr(S, Mid(ChrList, J, 1), Cmpr) Then HasChrList = True: Exit Function
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
