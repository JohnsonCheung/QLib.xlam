Attribute VB_Name = "MVb_PfxSfx"
Option Explicit

Function AddPfx$(A$, Pfx$)
AddPfx = Pfx & A
End Function

Function AddPfxSfx$(A$, Pfx$, Sfx$)
AddPfxSfx = Pfx & A & Sfx
End Function

Function AddSfx$(A$, Sfx$)
AddSfx = A & Sfx
End Function

Function AddPfxSpc_IfNonBlank$(A)
If A = "" Then Exit Function
AddPfxSpc_IfNonBlank = " " & A
End Function


Function AyAddPfx(A, Pfx) As String()
If Si(A) = 0 Then Exit Function
Dim O$(), U&, J&
U = UB(A)
ReDim O(U)
For J = 0 To U
    O(J) = Pfx & A(J)
Next
AyAddPfx = O
End Function

Function AyAddPfxSfx(A, Pfx, Sfx) As String()
Dim O$(), J&, U&
If Si(A) = 0 Then Exit Function
U = UB(A)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & A(J) & Sfx
Next
AyAddPfxSfx = O
End Function

Function AyAddSfx(A, Sfx) As String()
If Si(A) = 0 Then Exit Function
Dim O$(), J&, U&
U = UB(A)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = A(J) & Sfx
Next
AyAddSfx = O
End Function

Function AyIsAllEleHitPfx(A, Pfx$) As Boolean
If Si(A) = 0 Then Exit Function
Dim I
For Each I In A
   If Not HasPfx(I, Pfx) Then Exit Function
Next
AyIsAllEleHitPfx = True
End Function


Function AyAddCommaSpcSfxExlLas(Ay) As String()
Dim X, J, U%
U = UB(Ay)
For Each X In Itr(Ay)
    If J = U Then
        Push AyAddCommaSpcSfxExlLas, X
    Else
        Push AyAddCommaSpcSfxExlLas, X & ", "
    End If
    J = J + 1
Next
End Function
Function EnsSfx$(S, Sfx$)
If HasSfx(S, Sfx) Then
    EnsSfx = S
Else
    EnsSfx = S & Sfx
End If
End Function
Function TakSfxChr$(S, SfxChrLis$, Optional IsCasSen As Boolean)
If HasSfxChrLis(S, SfxChrLis, IsCasSen) Then TakSfxChr = LasChr(S)
End Function

Function HasSfxChrLis(A, SfxChrLis$, Optional IsCasSen As Boolean) As Boolean
Dim J%
For J = 1 To Len(SfxChrLis)
    If HasSfx(A, Mid(SfxChrLis, J, 1), IsCasSen) Then HasSfxChrLis = True: Exit Function
Next
End Function
Function HasPfx(S, Pfx, Optional IsCasSen As Boolean) As Boolean
HasPfx = StrComp(Left(S, Len(Pfx)), Pfx, IsCasSen) = 0
End Function
Function HasSfx(S, Sfx, Optional IsCasSen As Boolean) As Boolean
HasSfx = StrEq(Right(S, Len(Sfx)), Sfx, IsCasSen) = 0
End Function
Function StrEq(A, B, Optional IsCasSen) As Boolean
StrEq = StrComp(A, B, IIf(IsCasSen, VbCompareMethod.vbBinaryCompare, vbTextCompare))
End Function
Function HasSfxApIgnCas(A, ParamArray SfxAp()) As Boolean
Dim Av(): Av = SfxAp
HasSfxApIgnCas = HasSfxAv(A, Av)
End Function
Function HasSfxAp(A, ParamArray SfxAp()) As Boolean
Dim Av(): Av = SfxAp
HasSfxAp = HasSfxAv(A, Av, vbBinaryCompare)
End Function

Function HasSfxAv(A, SfxAv(), Optional IsCasSen As Boolean) As Boolean
Dim Sfx
For Each Sfx In SfxAv
    If HasSfx(A, Sfx, IsCasSen) Then HasSfxAv = True: Exit Function
Next
End Function

Function SyIsAllEleHitPfx(A$(), Pfx$) As Boolean
If Si(A) = 0 Then Exit Function
Dim I
For Each I In A
   If Not HasPfx(CStr(I), Pfx) Then Exit Function
Next
SyIsAllEleHitPfx = True
End Function

