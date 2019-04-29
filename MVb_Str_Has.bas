Attribute VB_Name = "MVb_Str_Has"
Option Explicit
Enum eIgnCas
    eeIgnCas = 0
    eeCasSen = 1
End Enum

Function HasDot(Str) As Boolean
HasDot = HasSubStr(Str, ".")
End Function

Function HasSubStr(Str, SubStr, Optional IgnCas As Boolean) As Boolean
If IgnCas Then
    HasSubStr = InStr(1, Str, SubStr, vbTextCompare) > 0
Else
    HasSubStr = InStr(1, Str, SubStr, vbBinaryCompare) > 0
End If
End Function

Function HasCrLf(A) As Boolean
HasCrLf = HasSubStr(A, vbCrLf)
End Function

Function HasHyphen(A) As Boolean
HasHyphen = HasSubStr(A, "-")
End Function

Function HasPound(A) As Boolean
HasPound = InStr(A, "#") > 0
End Function

Function HasSpc(S$) As Boolean
HasSpc = InStr(S, " ") > 0
End Function

Function HasSqBkt(S$) As Boolean
HasSqBkt = FstChr(S) = "[" And LasChr(S) = "]"
End Function

Function HasChrList(S$, ChrList$) As Boolean
Dim J%
For J = 1 To Len(ChrList)
    If HasSubStr(S, Mid(ChrList, J, 1)) Then HasChrList = True: Exit Function
Next
End Function

Function HasSubStrAy(S$, SubStrAy$()) As Boolean
Dim SubStr
For Each SubStr In SubStrAy
    If HasSubStr(S$, CStr(SubStr)) Then HasSubStrAy = True: Exit Function
Next
End Function
Function HasTT(S$, T1$, T2$) As Boolean
HasTT = Has2T(S, T1, T2)
End Function

Function HasVbar(S$) As Boolean
HasVbar = HasSubStr(S, "|")
End Function
