Attribute VB_Name = "QIde_Mth_Pfx"
Option Explicit
Private Const CMod$ = "MIde_Mth_Pfx."
Private Const Asm$ = "QIde"

Private Sub Z_MthPfx()
Ass MthPfx("Add_Cls") = "Add"
End Sub

Private Sub ZZ_MthPfx()
'Dim Ay$(): Ay = MthNyzV(CVbe)
'Dim Ay1$(): Ay1 = SyzMapAy(Ay, "MthPfx")
'ShwWs AyabWs(Ay, Ay1)
End Sub

Function MthPfxSyzMd(A As CodeModule) As String()
Dim N
For Each N In Itr(MthnyzMd(A))
    PushI MthPfxSyzMd, MthPfx(N)
Next
End Function

Function MthPfx$(Mthn)
Dim A0$
    A0 = Brk1(RmvPfxSy(Mthn, SplitVBar("ZZ_|Z_")), "__").S1
With Brk2(A0, "_")
    If .S1 <> "" Then
        MthPfx = .S1
        Exit Function
    End If
End With
Dim P2%
Dim Fnd As Boolean
    Dim C%
    Fnd = False
    For P2 = 2 To Len(A0)
        C = Asc(Mid(A0, P2, 1))
        If IsAscLCas(C) Then Fnd = True: Exit For
    Next
'---
    If Not Fnd Then Exit Function
Dim P3%
Fnd = False
    For P3 = P2 + 1 To Len(A0)
        C = Asc(Mid(A0, P3, 1))
        If IsAscUCas(C) Or IsAscDigit(C) Then Fnd = True: Exit For
    Next
'--
If Fnd Then
    MthPfx = Left(A0, P3 - 1)
    Exit Function
End If
MthPfx = Mthn
End Function


Private Sub ZZ()
Z_MthPfx
MIde_Mth_Pfx:
End Sub
