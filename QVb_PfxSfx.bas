Attribute VB_Name = "QVb_PfxSfx"
Option Explicit
Private Const CMod$ = "MVb_PfxSfx."
Private Const Asm$ = "QVb"

Function AddPfx$(S$, Pfx$)
AddPfx = Pfx & S
End Function

Function AddPfxSfx$(S$, Pfx$, Sfx$)
AddPfxSfx = Pfx & S & Sfx
End Function

Function AddSfx$(S$, Sfx$)
AddSfx = S & Sfx
End Function

Function AddPfxSpczIfNonBlank$(S$)
If S = "" Then Exit Function
AddPfxSpczIfNonBlank = " " & S
End Function


Function AddPfxzSy(Sy$(), Pfx$) As String()
Dim J&
For J = 0 To UB(Sy)
    PushI AddPfxzSy, Pfx & Sy(J)
Next
End Function

Function AddPfxzSySfx(Sy$(), Pfx$, Sfx$) As String()
Dim O$(), J&, U&
If Si(Sy) = 0 Then Exit Function
U = UB(Sy)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & Sy(J) & Sfx
Next
AddPfxzSySfx = O
End Function

Function AddSySfx(Sy$(), Sfx$) As String()
If Si(Sy) = 0 Then Exit Function
Dim O$(), J&, U&
U = UB(Sy)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Sy(J) & Sfx
Next
AddSySfx = O
End Function

Function IsAllEleHasPfx(Sy$(), Pfx$) As Boolean
Dim I
For Each I In Itr(Sy)
   If Not HasPfx(CStr(I), Pfx) Then Exit Function
Next
IsAllEleHasPfx = True
End Function

Function EnsSfx$(S$, Sfx$)
If HasSfx(S, Sfx) Then
    EnsSfx = S
Else
    EnsSfx = S & Sfx
End If
End Function
Function SfxChr$(S$, SfxChrLis$, Optional IsCasSen As Boolean)
If HasSfxChr(S, SfxChrLis, IsCasSen) Then SfxChr = LasChr(S)
End Function

Function HasSfxChr(S$, SfxChrLis$, Optional IsCasSen As Boolean) As Boolean
Dim J%
For J = 1 To Len(SfxChrLis)
    If HasSfx(S, Mid(SfxChrLis, J, 1), IsCasSen) Then HasSfxChr = True: Exit Function
Next
End Function
Function HasPfx(S$, Pfx$, Optional IsCasSen As Boolean) As Boolean
HasPfx = StrComp(Left(S, Len(Pfx)), Pfx, IsCasSen) = 0
End Function
Function HasSfx(S$, Sfx$, Optional IsCasSen As Boolean) As Boolean
HasSfx = IsStrEq(Right(S, Len(Sfx)), Sfx, IsCasSen) = 0
End Function
Function IsStrEq(A$, B$, Optional IsCasSen) As Boolean
IsStrEq = StrComp(A, B, IIf(IsCasSen, VbCompareMethod.vbBinaryCompare, vbTextCompare))
End Function
Function HasSfxApIgnCas(S$, ParamArray SfxAp()) As Boolean
Dim Av(): Av = SfxAp
HasSfxApIgnCas = HasSfxAv(S, Av)
End Function
Function HasSfxAp(S$, ParamArray SfxAp()) As Boolean
Dim Av(): Av = SfxAp
HasSfxAp = HasSfxAv(S, Av, vbBinaryCompare)
End Function

Function HasSfxAv(S$, SfxAv(), Optional IsCasSen As Boolean) As Boolean
Dim Sfx
For Each Sfx In SfxAv
    If HasSfx(S, CStr(Sfx), IsCasSen) Then HasSfxAv = True: Exit Function
Next
End Function

