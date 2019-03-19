Attribute VB_Name = "MVb_Str_Has"
Option Explicit
Enum eIgnCas
    eIgnCas = 0
    eCasSen = 1
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

Function HasSpc(A) As Boolean
HasSpc = InStr(A, " ") > 0
End Function

Function HasSqBkt(A) As Boolean
HasSqBkt = FstChr(A) = "[" And LasChr(A) = "]"
End Function

Function HasChrList(A, ChrList$) As Boolean
Dim J%
For J = 1 To Len(ChrList)
    If HasSubStr(A, Mid(ChrList, J, 1)) Then HasChrList = True: Exit Function
Next
End Function

Function HasSubStrAy(A, SubStrAy$()) As Boolean
Dim S
For Each S In SubStrAy
    If HasSubStr(A, CStr(S)) Then HasSubStrAy = True: Exit Function
Next
End Function
Function HasTT(L, T1, T2) As Boolean
If T1zLin(L) <> T1 Then Exit Function
If T2zLin(L) <> T2 Then Exit Function
HasTT = True
End Function

Function HasT1(L, T) As Boolean
HasT1 = T1(L) = T
End Function

Function HasVbar(A$) As Boolean
HasVbar = HasSubStr(A, "|")
End Function
