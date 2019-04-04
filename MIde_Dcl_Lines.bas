Attribute VB_Name = "MIde_Dcl_Lines"
Option Explicit
Public Const DocOfDclDic$ = "Key is PjNm.MdNm.  Value is Dcl (which is Lines)"
Public Const DocOfDcl$ = "It is Lines"
Private Sub Z_DclLinCnt()
Dim B1$(): B1 = CurSrc
Dim B2$(): B2 = SrcSrt(B1)
Dim A1%: A1 = DclLinCnt(B1)
Dim A2%: A2 = DclLinCnt(SrcSrt(B1))
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

Function DclLinCntzMd%(Md As CodeModule) 'Assume FstMth cannot have TopRmk
Dim I&
    I = FstMthLnoMd(Md)
    If I <= 0 Then
        DclLinCntzMd = Md.CountOfLines
        Exit Function
    End If
DclLinCntzMd = MthTopRmkLnoMdFm(Md, I) - 1
End Function

Function DclLinCnt%(Src$()) 'Assume FstMth cannot have TopRmk
Dim Top&
    Dim Fm&
    Fm = FstMthIx(Src)
    If Fm = -1 Then
        DclLinCnt = UB(Src) + 1
        Exit Function
    End If
    DclLinCnt = Fm
End Function

Function Dcl$(Src$())
Dcl = JnCrLf(DclLy(Src))
End Function

Function DclDicOfPj() As Dictionary
Set DclDicOfPj = DclDiczPj(CurPj)
End Function

Function DclDiczPj(A As VBProject) As Dictionary
If A.Protection = vbext_pp_locked Then Set DclDiczPj = New Dictionary: Exit Function
Dim C As VBComponent, M As CodeModule
Set DclDiczPj = New Dictionary
For Each C In A.VBComponents
    Set M = C.CodeModule
    Dim Dcl$: Dcl = DclzMd(M)
    If Dcl <> "" Then
        DclDiczPj.Add MdDNm(M), Dcl
    End If
Next
End Function

Function DclLy(Src$()) As String()
If Si(Src) = 0 Then Exit Function
Dim N&, O$()
   N = DclLinCnt(Src)
If N <= 0 Then Exit Function
O = AywFstNEle(Src, N)
DclLy = O
'Brw LyzNNAp("N Src DclLy", N, AyAddIxPfx(Src), O): Stop
End Function

Function DclzMd$(A As CodeModule)
Dim Cnt%
Cnt = DclLinCntzMd(A)
If Cnt = 0 Then Exit Function
DclzMd = LinesRmvBlankLinAtEnd(A.Lines(1, Cnt))
End Function

Function DclLyzMd(A As CodeModule) As String()
DclLyzMd = SplitCrLf(DclzMd(A))
End Function


Private Sub Z()
'Z_DclTyNm_TyLines
Z_DclLinCnt
MIde_Dcl_Lines:
End Sub
