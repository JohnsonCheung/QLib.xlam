Attribute VB_Name = "MVb_Str_Esc"
Option Explicit
Const CMod$ = "MVb_Str_Esc."
Function SlashCr$(S$)
SlashCr = SlashAsc$(S$, vbCr, "r")
End Function
Function EscOpnSqBkt$(S$)
EscOpnSqBkt = EscChr(S, "[")
End Function
Function EscClsSqBkt$(S$)
EscClsSqBkt = EscChr(S, "]")
End Function
Function UnSlashCrLf$(S$)
UnSlashCrLf = UnSlashCr(UnSlashLf(S))
End Function
Function SlashAsc$(S$, Asc%, C$)
SlashAsc = SlashChr(S, Chr(Asc), C)
End Function
Function SlashCrLf$(S$)
SlashCrLf = SlashLf(SlashCr(S))
End Function
Function SlashLf$(S$)
SlashLf = SlashAsc(S, vbLf, "n")
End Function
Function UnSlashChr$(S$, C$, SlashC$)
SlashChr = Replace(S, "\" & SlashC, C)
End Function

Function Slash$(S$, C$, SlashC$)
If InStr(S, "\" & SlashC) > 0 Then
    Debug.Print FmtQQ("SlashChr: Given S has \?, when UnSlash, it will not match", SlashC)
    Debug.Print vbTab; QuoteSq(S)
End If
SlashChr = Replace(S, C, "\" & SlashC)
End Function
Function UnEsc$(S$, C$)

End Function
Function UnEscBackSlash$(S$)
UnEscBackSlash = UnEsc(S, "\")
End Function
Function EscBackSlash$(S$)
EscBackSlash = EscChr(S, "\")
End Function

Function EscCr$(S$)
EscCr = EscAsc(S, vbCr)
End Function

Function EscCrLf$(S$)
EscCrLf = EscCr(EscLf(S))
End Function

Function EscLf$(S$)
EscLf = EscChr(S, Chr(vbLf))
End Function

Function SlashTab$(S$)
SlashTab = SlashChr(S, vbTab, "t")
End Function
Function EscAsc$(S$, A%)
End Function

Function UnSlashCr$(S$)
UnSlashCr = Replace(S, "\r", vbCr)
End Function
Function UnSlashLf$(S$)
UnSlashLf = Replace(S, "\n", vbLf)
End Function

Function TileSpc$(S$)
If InStr(S, "~") Then
    Debug.Print "TileSpc: Given-S has space"
    Debug.Print vbTab; "[" & S & "]"
End If
TileSpc = Replace(S, " ", "~")
End Function
Function UnTileSpc$(S$)
UnTileSpc = Replace(S, "~", " ")
End Function

Function UnSplashTab(S$)
EscUnTab = Replace(A, "\t", vbTab)
End Function

Function UnSlashBackSlash$(S$)
UnSlashBackSlash = Replace(A, "\\", "\")
End Function

Function UnEscCr$(S$)
UnEscCr = Replace(A, "\r", vbCr)
End Function

Function UnEscCrLf$(S$)
UnEscCrLf = UnEscLf(UnEscCr(S$))
End Function

Function UnEscLf$(S$)
UnEscLf = Replace(A, "\n", vbCr)
End Function

Function UnEscSpc$(S$)
UnEscSpc = Replace(A, "~", " ")
End Function

Function UnEscSqBkt$(S$)
UnEscSqBkt = Replace(A, Replace(A, "\o", "["), "\c", "]")
End Function

Function UnEscTab(S$)
UnEscTab = Replace(A, "\t", "~")
End Function
