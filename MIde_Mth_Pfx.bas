Attribute VB_Name = "MIde_Mth_Pfx"
Option Explicit

Private Sub Z_MthPfx()
Ass MthPfx("Add_Cls") = "Add"
End Sub

Private Sub ZZ_MthPfx()
'Dim Ay$(): Ay = MthNyzVbe(CurVbe)
'Dim Ay1$(): Ay1 = SyzAyMap(Ay, "MthPfx")
'WsVis AyabWs(Ay, Ay1)
End Sub

Function MthPfxAyMd(A As CodeModule) As String()
Dim N
For Each N In Itr(MthNyzMd(A))
    PushI MthPfxAyMd, MthPfx(N)
Next
End Function

Function MthPfx$(MthNm)
Dim A0$
    A0 = Brk1(RmvPfxAy(MthNm, SplitVBar("ZZ_|Z_")), "__").s1
With Brk2(A0, "_")
    If .s1 <> "" Then
        MthPfx = .s1
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
MthPfx = MthNm
End Function


Private Sub Z()
Z_MthPfx
MIde_Mth_Pfx:
End Sub
