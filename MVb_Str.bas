Attribute VB_Name = "MVb_Str"
Option Explicit
Function IsEqStr(A$, B$, Optional IgnoreCase As Boolean) As Boolean
IsEqStr = StrComp(A, B, IIf(IgnoreCase, vbTextCompare, vbBinaryCompare)) = 0
End Function

Function Pad0$(N%, NDig&)
Pad0 = Format(N, Dup("0", NDig))
End Function

Sub BrwStr(S$, Optional Fnn$, Optional UseVc As Boolean)
Dim T$: T = TmpFt("BrwStr", Fnn$)
WrtStr S, T
BrwFt T, UseVc
End Sub

Sub VcStr(S$, Optional Fnn$)
BrwStr S, Fnn, UseVc:=True
End Sub

Function StrDft$(S$, Dft$)
StrDft = IIf(S = "", Dft, S)
End Function

Function Dup$(S$, N&)
Dim O$, J&
For J = 0 To N - 1
    O = O & S
Next
Dup = O
End Function

Function HasSfxAs(S$, AsSfxAy$()) As Boolean
Dim I, Sfx$
For Each I In Itr(AsSfxAy)
    Sfx = I
    If HasSfx(S, Sfx) Then HasSfxAs = True: Exit Function
Next
End Function

Function HasPfxAs(S$, AsPfxAy$()) As Boolean
Dim I, Pfx$
For Each I In Itr(AsPfxAy)
    Pfx = I
    If HasPfx(S, Pfx) Then HasPfxAs = True: Exit Function
Next
End Function

Sub EdtStr(S$, Ft$)
WrtStr S, Ft, OvrWrt:=True
Brw Ft
End Sub
Function WrtStr$(S$, Ft$, Optional OvrWrt As Boolean)
If OvrWrt Then DltFfnIf Ft
Dim Fno%: Fno = FnoOup(Ft$)
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

