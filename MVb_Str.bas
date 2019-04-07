Attribute VB_Name = "MVb_Str"
Option Explicit
Function AddLib$(V, Lbl$)
Dim B$
If IsDate(V) Then
    B = FfnTimStr(CDate(V))
Else
    B = Replace(Replace(V, ";", "%3B"), "=", "%3D")
End If
If V <> "" Then AddLib = Lbl & "=" & B
End Function
Function IsEqStr(A, B, Optional IgnoreCase As Boolean) As Boolean
IsEqStr = StrComp(A, B, IIf(IgnoreCase, vbTextCompare, vbBinaryCompare)) = 0
End Function


Function Pad0$(N, NDig%)
Pad0 = Format(N, Dup("0", NDig))
End Function

Sub BrwStr(A, Optional Fnn$, Optional UseVc As Boolean)
Dim T$: T = TmpFt("BrwStr", Fnn$)
WrtStr A, T
BrwFt T, UseVc
End Sub

Sub VcStr(A, Optional Fnn$)
BrwStr A, Fnn, UseVc:=True
End Sub

Function StrDft$(A, B)
StrDft = IIf(A = "", B, A)
End Function

Function Dup$(S, N)
Dim O$, J&
For J = 0 To N - 1
    O = O & S
Next
Dup = O
End Function

Function HasStrSfxAy(A, SfxAy$()) As Boolean
Dim Sfx
For Each Sfx In Itr(SfxAy)
    If HasSfx(A, Sfx) Then HasStrSfxAy = True: Exit Function
Next
End Function

Function HasStrPfxAy(A, PfxAy$()) As Boolean
Dim Pfx
For Each Pfx In Itr(PfxAy)
    If HasPfx(A, Pfx) Then HasStrPfxAy = True: Exit Function
Next
End Function

Sub EdtStr(S, Ft)
WrtStr S, Ft, OvrWrt:=True
Brw Ft
End Sub
Function WrtStr$(Str, Ft, Optional OvrWrt As Boolean)
If OvrWrt Then DltFfnIf Ft
Dim Fno%: Fno = FnoOup(Ft)
Print #Fno, Str;
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
AddLib A, B
Pad0 A, C
BrwStr A, B
StrDft A, A
Dup A, A
HasStrSfxAy A, F
HasStrPfxAy A, F
WrtStr A, A, D
End Sub

