Attribute VB_Name = "MIde_Mth_Ix"
Option Explicit
Private Sub Z_MthIxAy()
Dim Ix, Src$()
Src = CurSrc
For Each Ix In MthIxItr(Src)
    If MthKd(Src(Ix)) <> "" Then
        Debug.Print Ix; Src(Ix)
    End If
Next
End Sub
Function MthIxItr(Src$(), Optional WhStr$)
Asg Itr(MthIxAy(Src, WhStr)), MthIxItr
End Function

Function EndLinIx&(Src$(), EndLinItm$, FmIx)
If 0 > FmIx Then EndLinIx = -1: Exit Function
Dim C$: C = "End " & EndLinItm
If HasSubStr(Src(FmIx), C) Then EndLinIx = FmIx: Exit Function
Dim O&
For O = FmIx + 1 To UB(Src)
   If HasPfx(Src(O), C) Then EndLinIx = O: Exit Function
Next
Thw CSub, "Cannot find EndLin", "EndLin FmIx Src", C, FmIx, Src
End Function
Function MthIx&(Src$(), MthNm)
Dim Ix
For Ix = 0 To UB(Src)
    If IsMthLinzNm(Src(Ix), MthNm) Then
        MthIx = Ix
        Exit Function
    End If
Next
MthIx = -1
End Function

Function MthIxAy(Src$(), Optional WhStr$) As Long()
Dim Ix
If WhStr = "" Then
    For Ix = 0 To UB(Src)
        If IsMthLin(Src(Ix)) Then
            PushI MthIxAy, Ix
        End If
    Next
Else
    Dim B As WhMth: Set B = WhMthzStr(WhStr)
    For Ix = 0 To UB(Src)
        If HitMthLin(Src(Ix), B) Then
            PushI MthIxAy, Ix
        End If
    Next
End If
End Function
Function MthIxzSrcNmTy(Src$(), MthNm, ShtMthTy$) As LngOpt
Dim Ix&
For Ix = 0 To UB(Src)
    With MthNm3(Src(Ix))
        If .Nm = MthNm Then
            If .ShtTy = ShtMthTy Then
                MthIxzSrcNmTy = SomLng(CLng(Ix))
                Exit Function
            End If
        End If
    End With
Next
Stop
End Function

Function MthIxAyzNm(Src$(), MthNm$) As Long()
Dim Ix&
Ix = MthIxzFst(Src, MthNm)
If Ix = -1 Then Exit Function
PushI MthIxAyzNm, Ix
If IsPrpLin(Src(Ix)) Then
    Ix = MthIxzFst(Src, MthNm, Ix + 1)
    If Ix > 0 Then Push MthIxAyzNm, Ix
End If
End Function

Function MthIxzFst&(Src$(), MthNm, Optional SrcFmIx& = 0)
Dim I
For I = SrcFmIx To UB(Src)
    If MthNmzLin(Src(I)) = MthNm Then
        MthIxzFst = I
        Exit Function
    End If
Next
MthIxzFst = -1
End Function

Function MthToIxAy(Src$(), FmIxAy&()) As Long()
Dim Ix
For Each Ix In Itr(FmIxAy)
    PushI MthToIxAy, MthToIx(Src, Ix)
Next
End Function

Function MthToIx&(Src$(), MthIx)
MthToIx = EndLinIx(Src, MthKd(Src(MthIx)), MthIx)
End Function

Function FstMthLnozMd&(Md As CodeModule)
Dim J&
For J = 1 To Md.CountOfLines
   If IsMthLin(Md.Lines(J, 1)) Then
       FstMthLnozMd = J
       Exit Function
   End If
Next
End Function

Function FstMthIx&(Src$())
Dim J&
For J = 0 To UB(Src)
   If IsMthLin(Src(J)) Then
       FstMthIx = J
       Exit Function
   End If
Next
FstMthIx = -1
End Function

Function MthLnoMdMth&(A As CodeModule, MthNm$)
MthLnoMdMth = 1 + MthIxzFst(Src(A), MthNm, 0)
End Function

Function MthLnoAyMdMth(A As CodeModule, MthNm$) As Long()
MthLnoAyMdMth = AyIncEle1(MthIxAyzNm(Src(A), MthNm))
End Function

Private Sub Z()
MIde_Mth_Ix:
End Sub

Function MthRgAy(Src$()) As MthRg()
If Si(Src) = 0 Then Exit Function
Dim F&(), T&(), N$()
F = MthIxAy(Src)
N = MthNyzSrcFm(Src, F)
T = MthToIxAy(Src, F)
Dim S&
S = Si(F)
If S = 0 Then Exit Function
Dim O() As MthRg
ReDim O(S - 1)
Dim J&
For J = 0 To S - 1
    Set O(J) = New MthRg
    With O(J)
        .MthNm = N(J)
        .FmIx = F(J)
        .ToIx = T(J)
    End With
Next
MthRgAy = O
End Function

