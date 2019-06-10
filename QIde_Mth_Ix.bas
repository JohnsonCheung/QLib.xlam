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
Function VbaItm$(Lin)
Dim O$: O = T1(RmvMdy(Lin))
If IsVbaItm(O) Then VbaItm = O
End Function
Function IsVbaItm(Itm$) As Boolean
IsVbaItm = HasEle(VbaItmAy, Itm)
End Function
Function VbaItmAyV() As String()
VbaItmAyV = VbaItmAyzSrc(SrcV)
End Function

Function VbaItmAyzSrc(Src$()) As String()
Dim S
For Each S In Itr(Src)
    PushNonBlank VbaItmAyzSrc, VbaItm(S)
Next
End Function
Function VbaItmAy() As String()
Static X As Boolean, Y
If Not X Then
    X = True
    Y = SyzSS("Function Sub Type Enum Property")
End If
VbaItmAy = Y
End Function
Function EndLinzVbaItm$(Itm$)
If Not IsVbaItm(Itm) Then Thw CSub, "Given Itm is not a VbaItm", "Itm", Itm
EndLinzVbaItm = "End " & Itm
End Function

Function EndLin$(Src$(), ItmIx)
EndLin = EndLinzVbaItm(VbaItm(Src(ItmIx)))
End Function

Function EndLinzM$(M As CodeModule, ItmLno)
EndLinzM = EndLinzVbaItm(VbaItm(M.Lines(ItmLno, 1)))
End Function

Function EndLnozM&(M As CodeModule, ItmLno)
Dim EndL$, O&
EndL = EndLinzM(M, ItmLno)
If HasSubStr(M.Lines(ItmLno, 1), EndL) Then EndLnozM = ItmLno: Exit Function
For O = ItmLno + 1 To M.CountOfLines
   If HasPfx(M.Lines(O, 1), EndL) Then EndLnozM = O: Exit Function
Next
Thw CSub, "Cannot find EndLin", "EndLin ItmLno Md", EndL, ItmLno, Mdn(M)
End Function

Function EndLix&(Src$(), ItmIx)
Dim EndL$, O&
EndL = EndLin(Src, ItmIx)
If HasSubStr(Src(ItmIx), EndL) Then EndLix = ItmIx: Exit Function
For O = ItmIx + 1 To UB(Src)
   If HasPfx(Src(O), EndL) Then EndLix = O: Exit Function
Next
Thw CSub, "Cannot find EndLin", "EndLin ItmIx Src", EndL, ItmIx, Src
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

Function MthIxyzMN(M As CodeModule, Mthn) As Long()
MthIxyzMN = MthIxyzSN(Src(M), Mthn)
End Function

Function MthIxzMTN&(M As CodeModule, ShtMthTy$, Mthn)
MthIxzMTN = MthIxzSTN(Src(M), ShtMthTy, Mthn)
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
Ix = FstMthIxzN(Src, Mthn)
If Ix = -1 Then Exit Function
PushI MthIxyzSN, Ix
If IsPrpLin(Src(Ix)) Then
    Ix = FstMthIxzN(Src, Mthn, Ix + 1)
    If Ix > 0 Then PushI MthIxyzSN, Ix
End If
End Function
Function FstMthIx&(Src$())
Dim O&
Dim L
For Each L In Itr(Src)
    If IsMthLin(L) Then FstMthIx = O: Exit Function
Next
FstMthIx = -1
End Function

Function FstMthIxzN&(Src$(), Mthn, Optional SrcFmIx& = 0)
Dim I
For I = SrcFmIx To UB(Src)
    If MthnzLin(Src(I)) = Mthn Then
        FstMthIxzN = I
        Exit Function
    End If
Next
FstMthIxzN = -1
End Function

Function MthEIxy(Src$(), FmIxy&()) As Long()
Dim Ix, ELin$
Stop
For Each Ix In Itr(FmIxy)
    PushI MthEIxy, EndLix(Src, Ix)
Next
End Function

Function MthFeizSIW(Src$(), MthIx) As Fei
MthFeizSIW = Fei(MthIx, EndLix(Src, MthIx))
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

Function MthLnozMM&(M As CodeModule, Mthn)
MthLnozMM = 1 + FstMthIxzN(Src(M), Mthn, 0)
End Function


Function MthLnoAyzMN(M As CodeModule, Mthn) As Long()
MthLnoAyzMN = AyIncEle1(MthIxyzSN(Src(M), Mthn))
End Function

Private Sub ZZZ()
MIde_Mth_Ix:
End Sub


