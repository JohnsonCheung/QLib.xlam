Attribute VB_Name = "QIde_Mth_Ix"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mth_Ix."
Private Const Asm$ = "QIde"
Private Sub Z_MthIxy()
Dim Ix, Src$()
Src = CSrc
For Each Ix In MthIxItr(Src)
    If MthKd(Src(Ix)) <> "" Then
        Debug.Print Ix; Src(Ix)
    End If
Next
End Sub
Function MthIxItr(Src$())
Asg Itr(MthIxy(Src)), MthIxItr
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

Function MthIxy(Src$()) As Long()
Dim Ix
For Ix = 0 To UB(Src)
    If IsMthLin(Src(Ix)) Then
        PushI MthIxy, Ix
    End If
Next
End Function
Function FstMthIxzSN&(Src$(), Mthn)
Dim Ix&
For Ix = 0 To UB(Src)
    With Mthn3zL(Src(Ix))
        If .Nm = Mthn Then
            FstMthIxzSN = Ix
            Exit Function
        End If
    End With
Next
FstMthIxzSN = -1
End Function

Function MthIxyzMN(A As CodeModule, Mthn) As Long()
MthIxyzMN = MthIxyzSN(Src(A), Mthn)
End Function

Function MthIxzMTN&(A As CodeModule, ShtMthTy$, Mthn)
MthIxzMTN = MthIxzSTN(Src(A), ShtMthTy, Mthn)
End Function

Function MthIxzSTN&(Src$(), ShtMthTy$, Mthn)
Dim Ix&
For Ix = 0 To UB(Src)
    With Mthn3zL(Src(Ix))
        If .Nm = Mthn Then
            If .ShtTy = ShtMthTy Then
                MthIxzSTN = Ix
                Exit Function
            End If
        End If
    End With
Next
MthIxzSTN = -1
End Function

Function MthIxyzSN(Src$(), Mthn) As Long()
Dim Ix&
Ix = FstMthIx(Src, Mthn)
If Ix = -1 Then Exit Function
PushI MthIxyzSN, Ix
If IsPrpLin(Src(Ix)) Then
    Ix = FstMthIx(Src, Mthn, Ix + 1)
    If Ix > 0 Then PushI MthIxyzSN, Ix
End If
End Function

Function FstMthIx&(Src$(), Mthn, Optional SrcFmIx& = 0)
Dim I
For I = SrcFmIx To UB(Src)
    If MthnzLin(Src(I)) = Mthn Then
        FstMthIx = I
        Exit Function
    End If
Next
FstMthIx = -1
End Function

Function MthEIxy(Src$(), FmIxy&()) As Long()
Dim Ix
For Each Ix In Itr(FmIxy)
    PushI MthEIxy, MthEIx(Src, Ix)
Next
End Function
Function MthIxzSIW&(Src$(), MthIx, WiTopRmk As Boolean)
If WiTopRmk Then
    MthIxzSIW = TopRmkIx(Src, MthIx)
Else
    MthIxzSIW = MthIx
End If
End Function

Function MthFeizSIW(Src$(), MthIx, Optional WiTopRmk As Boolean) As Fei
MthFeizSIW = Fei(MthIxzSIW(Src, MthIx, WiTopRmk), MthEIx(Src, MthIx))
End Function

Function MthEIx&(Src$(), MthIx)
MthEIx = EndLinIx(Src, MthKd(Src(MthIx)), MthIx)
End Function

Function FstMthLnozM&(Md As CodeModule)
Dim J&
For J = 1 To Md.CountOfLines
   If IsMthLin(Md.Lines(J, 1)) Then
       FstMthLnozM = J
       Exit Function
   End If
Next
End Function

Function FstMthIxzS&(Src$())
Dim J&
For J = 0 To UB(Src)
   If IsMthLin(Src(J)) Then
       FstMthIxzS = J
       Exit Function
   End If
Next
FstMthIxzS = -1
End Function

Function MthLnoMdMth&(A As CodeModule, Mthn)
MthLnoMdMth = 1 + FstMthIx(Src(A), Mthn, 0)
End Function

Function MthLnoAyzMN(A As CodeModule, Mthn) As Long()
MthLnoAyzMN = AyIncEle1(MthIxyzSN(Src(A), Mthn))
End Function

Private Sub ZZZ()
MIde_Mth_Ix:
End Sub


