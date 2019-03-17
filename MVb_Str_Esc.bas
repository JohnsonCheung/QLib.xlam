Attribute VB_Name = "MVb_Str_Esc"
Option Explicit
Const CMod$ = "MVb_Str_Esc."
Function Esc$(A, Fm$, ToStr$)
Const CSub$ = CMod & "Esc"
If InStr(A, ToStr) > 0 Then
    Inf CSub, "Warning: escaping a {Str} of {FmStrSub} to {ToSubStr} is found that {Str} contains some {ToSubStr}.  This will make the string chagned after EscUn", A, Fm, ToStr
End If
Esc = Replace(A, Fm, ToStr)
End Function

Function EscBackSlash$(A)
EscBackSlash = Replace(A, "\", "\\")
End Function

Function EscCr$(A)
EscCr = Esc(A, vbCr, "\r")
End Function

Function EscCrLf$(A)
EscCrLf = EscCr(EscLf(A))
End Function

Function EscKey$(A)
EscKey = EscCrLf(EscSpc(EscTab(A)))
End Function

Function EscLf$(A)
EscLf = Esc(A, vbLf, "\n")
End Function

Function EscSpc$(A)
EscSpc = Esc(A, " ", "~")
End Function

Function EscSqBkt$(A)
EscSqBkt = Replace(Replace(A, "[", "\o"), "]", "\c")
End Function

Function EscTab$(A)
EscTab = Esc(A, vbTab, "\t")
End Function

Function EscUnCr$(A)
EscUnCr = Replace(A, "\r", vbCr)
End Function

Function EscUnSpc$(A)
EscUnSpc = Replace(A, "~", " ")
End Function

Function EscUnTab(A)
EscUnTab = Replace(A, "\t", "~")
End Function

Function UnEscBackSlash$(A)
UnEscBackSlash = Replace(A, "\\", "\")
End Function

Function UnEscCr$(A)
UnEscCr = Replace(A, "\r", vbCr)
End Function

Function UnEscCrLf$(A)
UnEscCrLf = UnEscLf(UnEscCr(A))
End Function

Function UnEscLf$(A)
UnEscLf = Replace(A, "\n", vbCr)
End Function

Function UnEscSpc$(A)
UnEscSpc = Replace(A, "~", " ")
End Function

Function UnEscSqBkt$(A)
UnEscSqBkt = Replace(A, Replace(A, "\o", "["), "\c", "]")
End Function

Function UnEscTab(A)
UnEscTab = Replace(A, "\t", "~")
End Function
