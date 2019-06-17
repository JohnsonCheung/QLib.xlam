Attribute VB_Name = "QVb_PfxSfx"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_PfxSfx."
Private Const Asm$ = "QVb"

Function SzAddP(S, Pfx)
SzAddP = Pfx & S
End Function

Function SzAddPS(S, Pfx, Sfx)
SzAddPS = Pfx & S & Sfx
End Function
Function IsNB(S) As Boolean
IsNB = Trim(S) <> ""
End Function
Function SzAddSIfNB$(SIfNB, Sfx)
If IsNB(SIfNB) Then SzAddSIfNB = SIfNB & Sfx
End Function
Function AddSfx(S, Sfx)
AddSfx = S & Sfx
End Function

Function SzAddPSpczIfNonBlank$(S)
If S = "" Then Exit Function
SzAddPSpczIfNonBlank = " " & S
End Function

Function SyzTrim(Ay) As String()
Dim V
For Each V In Itr(Ay)
    PushI SyzTrim, Trim(V)
Next
End Function

Function SyzAyP(Ay, Pfx) As String()
Dim I
For Each I In Itr(Ay)
    PushI SyzAyP, Pfx & I
Next
End Function

Function SyzAyPS(Ay, Pfx, Sfx) As String()
Dim I
For Each I In Itr(Ay)
    PushI SyzAyPS, Pfx & I & Sfx
Next
End Function

Function SzAddSIfNBzAy(Ay, Sfx) As String()
Dim I, S$
For Each I In Itr(Ay)
    PushI SzAddSIfNBzAy, SzAddSIfNB(I, Sfx)
Next
End Function

Function SyzAyS(Ay, Sfx) As String()
Dim I
For Each I In Itr(Ay)
    PushI SyzAyS, I & Sfx
Next
End Function

Function IsSyzAllEleHasPfx(Sy$(), Pfx) As Boolean
Dim I
For Each I In Itr(Sy)
   If Not HasPfx(I, Pfx) Then Exit Function
Next
IsSyzAllEleHasPfx = True
End Function

Function EnsSfx(S, Sfx)
If HasSfx(S, Sfx) Then
    EnsSfx = S
Else
    EnsSfx = S & Sfx
End If
End Function
Function SfxChr$(S, SfxChrLis$, Optional C As VbCompareMethod = vbBinaryCompare)
If HasSfxChr(S, SfxChrLis, C) Then SfxChr = LasChr(S)
End Function

Function HasSfxChr(S, SfxChrLis$, Optional C As VbCompareMethod = vbBinaryCompare) As Boolean
Dim J%
For J = 1 To Len(SfxChrLis)
    If HasSfx(S, Mid(SfxChrLis, J, 1), C) Then HasSfxChr = True: Exit Function
Next
End Function
Function HasPfxOfAllEle(Ay, Pfx, Optional C As VbCompareMethod = vbTextCompare) As Boolean
If Si(Ay) = 0 Then Exit Function
Dim V
For Each V In Itr(Ay)
    If Not HasPfx(V, Pfx, C) Then Exit Function
Next
HasPfxOfAllEle = True
End Function
Function HasPfx(S, Pfx, Optional C As VbCompareMethod = vbTextCompare) As Boolean
HasPfx = StrComp(Left(S, Len(Pfx)), Pfx, C) = 0
End Function
Function HasPfxss(S, Pfxss$, Optional C As VbCompareMethod = vbTextCompare) As Boolean
End Function
Function HasPfxAy(S, PfxAy$(), Optional C As VbCompareMethod = vbTextCompare) As Boolean
Dim Pfx: For Each Pfx In Itr(PfxAy)
    If HasPfx(S, Pfx, C) Then HasPfxAy = True: Exit Function
Next
End Function

Function HasPfxzAy(Ay, Pfx, Optional C As VbCompareMethod = vbTextCompare) As Boolean
Dim I
For Each I In Itr(Ay)
    If HasPfx(I, Pfx, C) Then HasPfxzAy = True: Exit Function
Next
End Function
Function HasSfx(S, Sfx, Optional C As VbCompareMethod = vbBinaryCompare) As Boolean
HasSfx = IsEqStr(Right(S, Len(Sfx)), Sfx, C)
End Function
Function HasSfxApIgnCas(S, ParamArray SfxAp()) As Boolean
Dim Av(): Av = SfxAp
HasSfxApIgnCas = HasSfxAv(S, Av, vbTextCompare)
End Function
Function HasSfxApCasSen(S, ParamArray SfxAp()) As Boolean
Dim Av(): Av = SfxAp
HasSfxApCasSen = HasSfxAv(S, Av, vbBinaryCompare)
End Function

Function HasSfxAv(S, SfxAv(), C As VbCompareMethod) As Boolean
Dim Sfx
For Each Sfx In SfxAv
    If HasSfx(S, Sfx, C) Then HasSfxAv = True: Exit Function
Next
End Function

