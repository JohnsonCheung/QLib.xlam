Attribute VB_Name = "MIde_Loc_Jmp"
Option Explicit
Sub MdJmpLno(A As CodeModule, Lno&)
'MdJmpRRCC A, RRCC(Lno, Lno, 1, 1)
End Sub
Sub Jmp(MdNm$)
JmpMdNm MdNm
End Sub
Sub JmpMdNm(MdNm$)
JmpMd Md(MdNm)
End Sub
Sub JmpPos(A As MdPos)
JmpMd A.Md
JmpLinPos A.Pos
End Sub

Sub LinPosAsg(A As LinPos, OLno&, OC1%, OC2%)
With A
    OLno = .Lno
    OC1 = .Pos.Cno1
    OC2 = .Pos.Cno2
End With
End Sub
Sub JmpLinPos(A As LinPos)
Dim L&, C1%, C2%
LinPosAsg A, L, C1, C2
With CurCdPne
    .TopLine = L
    .SetSelection L, C1, L, C2
End With
SendKeys "^{F4}"
End Sub

Sub JmpLno(Lno&)
JmpLinPos LinPosLno(Lno)
End Sub
Function MdLinPos(A As CodeModule, Lno&) As MdPos
MdLinPos = MdPos(A, LinPos(Lno, EmpPos))
End Function
Function EmpPos() As Pos
End Function
Sub JmpMdMth(A As CodeModule, MthNm$)
JmpMd A
JmpMth MthNm
End Sub
Function LinPosMth(MthNm$) As LinPos
Dim MthLin$, Src$(), FmIx&
Src = SrcMd
FmIx = FstMthIxzMth(Src, MthNm)
If FmIx = -1 Then Exit Function
MthLin = Src(FmIx)
LinPosMth = LinPos(FmIx + 1, MthPos(MthLin))
End Function
Sub JmpMth(MthNm$)
JmpLinPos LinPosMth(MthNm)
End Sub
Sub JmpCurMd()
JmpMd CurMd
End Sub
Sub JmpMd(A As CodeModule)
JmpCmp A.Parent.Name
End Sub
Sub JmpPj(A As VBProject)
ClsWin
Dim Md As CodeModule
Set Md = FstMd(A)
If IsNothing(Md) Then
    Exit Sub
End If
Md.CodePane.Show
TileVBtn.Execute
DoEvents
End Sub

Sub JmpPjMd(A As VBProject, MdNm$)
ClsWinzExl WinzMd(Md(MdNm))
End Sub

Sub MdJmp(A As CodeModule)
A.CodePane.Show
End Sub

Sub JmpClsNm(ClsNm$)
End Sub


Sub JmpCmp(CmpNm$)
ClsWinzExptMdImm CmpNm
ShwCmp CmpNm
TileV
End Sub


