Attribute VB_Name = "QVb_Str_PfxSfx"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_PfxSfx."
Private Const Asm$ = "QVb"

Function AddPfx(S, Pfx)
AddPfx = Pfx & S
End Function

Function AddPfxS(S, Pfx, Sfx)
AddPfxS = Pfx & S & Sfx
End Function
Function IsNB(S) As Boolean
IsNB = Trim(S) <> ""
End Function

Function AddNBSfx$(S, Sfx)
If IsNB(S) Then AddNBSfx = S & Sfx
End Function
Function AddNBPfx$(S, Pfx)
If IsNB(S) Then AddNBPfx = Pfx & S
End Function
Function AddSfx(S, Sfx)
AddSfx = S & Sfx
End Function

Function AddPfxSpczIfNB$(S)
If S = "" Then Exit Function
AddPfxSpczIfNB = " " & S
End Function

Function SyzTrim(Ay) As String()
Dim V
For Each V In Itr(Ay)
    PushI SyzTrim, Trim(V)
Next
End Function

Function AddPfxzAy(Ay, Pfx) As String()
Dim I
For Each I In Itr(Ay)
    PushI AddPfxzAy, Pfx & I
Next
End Function

Function AddPfxSzAy(Ay, Pfx, Sfx) As String()
Dim I
For Each I In Itr(Ay)
    PushI AddPfxSzAy, Pfx & I & Sfx
Next
End Function

Function AddNBSfxzAy(Ay, Sfx) As String()
Dim I, S$
For Each I In Itr(Ay)
    PushI AddNBSfxzAy, AddNBSfx(I, Sfx)
Next
End Function

Function AddSfxzAy(Ay, Sfx) As String()
Dim I
For Each I In Itr(Ay)
    PushI AddSfxzAy, I & Sfx
Next
End Function

Function IsSyzAllHasPfxzSomEle(Sy$(), Pfx) As Boolean
Dim I
For Each I In Itr(Sy)
   If Not HasPfx(I, Pfx) Then Exit Function
Next
IsSyzAllHasPfxzSomEle = True
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

Function Sfx$(S, Suffix$, Optional C As VbCompareMethod = vbBinaryCompare)
If HasSfx(S, Suffix, C) Then Sfx = Suffix
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
Dim PfxAy$(): PfxAy = SyzSS(Pfxss)
HasPfxss = HasPfxAy(S, PfxAy, C)
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


Function PfxSySpc$(S, PfxSy$()) ' Return Fst ele-P of [PfxSy] if [S] has pfx ele-P and a space
Dim P
For Each P In PfxSy
    If HasPfx(S, P & " ") Then PfxSySpc = P: Exit Function
Next
End Function

Function PfxzAy$(S, PfxSy$())
'Ret : #PfxAy-Space ! fst ele-$ of @PfxSy so that @S has $.
Dim P
For Each P In PfxSy
    If HasPfx(S, P) Then PfxzAy = P: Exit Function
Next
End Function

Function PfxzAyS$(S, PfxAy$())
'Ret : #Pfx-Ay-Spc# ! fst pfx-$ in @PfxAy so that @S &HasPfx $ & " "
Dim P: For Each P In PfxAy
    If HasPfx(S, P & " ") Then PfxzAyS = P: Exit Function
Next
End Function

Function SfxzAyS$(S, SfxAy$())
'Ret : #Sfx-Ay-Spc# ! fst ele-$ of @SfxAy if @] has pfx $.
Dim Sfx: For Each Sfx In SfxAy
    If HasSfx(S, Sfx) Then SfxzAyS = Sfx: Exit Function
Next
End Function

Function PfxS$(S, P$)
'Ret: #Sfx-Space ! @Pfx if @S has has such @Pfx+" " else ""
PfxS = Pfx(S, P & " ")
End Function

Function Pfx$(S, P$)
'Ret: #Sfx-Space ! @Pfx if @S has has such @Pfx+" " else ""
If HasPfx(S, P) Then Pfx = P
End Function

Function PfxzPfxAp(S, ParamArray PfxAp())
Dim PfxSy$(): PfxSy = SyzAy(PfxSy)
PfxzPfxAp = PfxzPfxSy(S, PfxSy)
End Function


