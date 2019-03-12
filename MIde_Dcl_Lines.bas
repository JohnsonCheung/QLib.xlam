Attribute VB_Name = "MIde_Dcl_Lines"
Option Explicit
Private Sub Z_DclLinCnt()
Dim B1$(): B1 = SrcMd
Dim B2$(): B2 = SrtedSrc(B1)
Dim A1%: A1 = DclLinCnt(B1)
Dim A2%: A2 = DclLinCnt(SrtedSrc(B1))
End Sub

Sub BrwDclLinCntDryPj()
BrwDry DclLinCntDryzPj(CurPj)
End Sub

Function DclLinCntDryzPj(A As VBProject) As Variant()
Dim C As VBComponent
For Each C In A.VBComponents
    PushI DclLinCntDryzPj, Array(C.Name, DclLinCntzMd(C.CodeModule))
Next
End Function

Function DclLinCntzMd%(Md As CodeModule)
Dim I&
    I = FstMthLnoMd(Md)
    If I <= 0 Then
        DclLinCntzMd = Md.CountOfLines
        Exit Function
    End If
    I = MthTopRmkLnoMdFm(Md, I)
Dim O&
    For I = I - 1 To 1 Step -1
         If IsCdLin(Md.Lines(I, 1)) Then O = I + 1: GoTo X
    Next
    O = 0
X:
DclLinCntzMd = O
End Function

Function DclLinCnt%(Src$())
Dim Top&
    Dim Fm&
    Fm = FstMthIx(Src)
    If Fm = -1 Then
        DclLinCnt = UB(Src) + 1
        Exit Function
    End If
    Top = MthTopRmkIx(Src, Fm)
    If Top = -1 Then
        Top = Fm
    End If
Dim O&
    Dim I&
    For I = Top - 1 To 0 Step -1
         If IsCdLin(Src(I)) Then O = I + 1: GoTo X
    Next
    O = 0
X:
DclLinCnt = O
End Function

Function DclLines$(Src$())
DclLines = JnCrLf(DclLy(Src))
End Function

Function DclLy(Src$()) As String()
If Sz(Src) = 0 Then Exit Function
Dim N&
   N = DclLinCnt(Src)
If N = 0 Then Exit Function
DclLy = AywFstNEle(Src, N)
End Function

Function DclLineszMd$(A As CodeModule)
Dim Cnt%
Cnt = DclLinCntzMd(A)
If Cnt = 0 Then Exit Function
DclLineszMd = A.Lines(1, Cnt)
End Function

Function DclLyzMd(A As CodeModule) As String()
DclLyzMd = SplitCrLf(DclLineszMd(A))
End Function


Private Sub Z()
'Z_DclTyNm_TyLines
Z_DclLinCnt
MIde_Dcl_Lines:
End Sub
