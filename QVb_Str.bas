Attribute VB_Name = "QVb_Str"
Option Explicit
Private Const CMod$ = "MVb_Str."
Private Const Asm$ = "QVb"

Function SzIf$(IfTrue As Boolean, S)
If IfTrue Then SzIf = S
End Function

Function IsEqStr(A, B, Optional C As VbCompareMethod = vbBinaryCompare) As Boolean
IsEqStr = StrComp(A, B, C) = 0
End Function

Function Pad0$(N&, NDig&)
Pad0 = Format(N, Dup("0", NDig))
End Function

Sub BrwStr(S, Optional Fnn$, Optional UseVc As Boolean)
Dim T$: T = TmpFt("BrwStr", Fnn$)
WrtStr S, T
BrwFt T, UseVc
End Sub

Sub VcStr(S, Optional Fnn$)
BrwStr S, Fnn, UseVc:=True
End Sub
Function StrDft$(S, Dft)
StrDft = IIf(S = "", Dft, S)
End Function

Function Dup$(S, N&)
Dim O$, J&
For J = 0 To N - 1
    O = O & S
Next
Dup = O
End Function

Function HasSfxAs(S, AsSfxSy$()) As Boolean
Dim I, Sfx$
For Each I In Itr(AsSfxSy)
    Sfx = I
    If HasSfx(S, Sfx) Then HasSfxAs = True: Exit Function
Next
End Function

Function HasPfxAs(S, AsPfxSy$()) As Boolean
Dim I, Pfx$
For Each I In Itr(AsPfxSy)
    Pfx = I
    If HasPfx(S, Pfx) Then HasPfxAs = True: Exit Function
Next
End Function

Sub EdtStr(S, Ft)
WrtStr S, Ft, OvrWrt:=True
Brw Ft
End Sub
Function WrtStr$(S, Ft, Optional OvrWrt As Boolean)
If OvrWrt Then DltFfnIf Ft
Dim Fno%: Fno = FnoO(Ft)
Print #Fno, S;
Close #Fno
WrtStr = Ft
End Function


Private Sub ZZ()
Dim A As Variant
Dim B$
Dim C%
Dim D As Boolean
Dim E&
Dim F$()
End Sub

