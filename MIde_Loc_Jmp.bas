Attribute VB_Name = "MIde_Loc_Jmp"
Option Explicit
Sub MdJmpLno(A As CodeModule, Lno&)
JmpMd A
JmpLno Lno
End Sub
Function HasMdNm(MdNm$, Optional Inf As Boolean) As Boolean
If HasItn(CurPj.VBComponents, MdNm) Then HasMdNm = True: Exit Function
If Inf Then InfLin CSub, FmtQQ("Md[?] not exist", MdNm)
End Function
Sub Jmp(MdNmDotLno$)
JmpMdNm BefOrAll(MdNmDotLno, ".")
JmpLno Aft(MdNmDotLno, ".")
End Sub
Function JmpMdNm%(MdNm$)
If Not HasMdNm(MdNm, Inf:=True) Then JmpMdNm = 1: Exit Function
JmpMd Md(MdNm)
TileV
End Function
Sub JmpMdPos(A As MdPos)
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

Sub JmpLin(Lno&)
With CurCdPne
    .TopLine = Lno
    .SetSelection Lno, 1, Lno, Len(CurMd.Lines(Lno, 1))
End With

End Sub

Sub JmpLinPos(A As LinPos)
Dim L&, C1%, C2%
LinPosAsg A, L, C1, C2
With CurCdPne
    If C1 = 0 Or C2 = 0 Then
        C1 = 1
        C2 = Len(.CodeModule.Lines(L, 1)) + 1
    End If
    .TopLine = L
    .SetSelection L, C1, L, C2
End With
'SendKeys "^{F4}"
End Sub

Sub JmpLno(Lno&)
JmpLinPos LinPos(Lno)
End Sub
Function MdPoszLno(A As CodeModule, Lno&) As MdPos
MdPoszLno = MdPos(A, LinPos(Lno))
End Function
Function EmpPos() As Pos
Static O As New Pos
Set EmpPos = O
End Function
Sub JmpMdMth(A As CodeModule, MthNm$)
JmpMd A
JmpMth MthNm
End Sub
Function MdPosAyzMth(MthNm$) As MdPos()
Dim MdNm
For Each MdNm In Itr(MdNyzMth(MthNm))
Dim MthLin$, Src$(), FmIx&
Src = CurSrc
FmIx = MthIxzFst(Src, MthNm)
If FmIx = -1 Then Exit Function
MthLin = Src(FmIx)
MdPosAyzMth = LinPos(FmIx + 1, MthPos(MthLin))
Next
End Function

Sub JmpMth(MthNm$)
Dim A() As MdPos: A = MdPosAyzMth(MthNm)
Stop
Select Case Si(A)
Case 0: Debug.Print FmtQQ("Mth[?] not found", MthNm)
Case 1: JmpLinPos A(0)
Case 2:
    Dim I
    For Each I In A
        Debug.Print "JmpLinPos " & LinPosStr(CvLinPos(I))
    Next
End Select
End Sub

Function LinPosStr$(A As LinPos)
With A
LinPosStr = FmtQQ("Lin ? C1 ? C2 ?", .Lno, .Pos.Cno1, .Pos.Cno1)
End With
End Function
Sub JmpCurMd()
JmpMd CurMd
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

Sub JmpMd(A As CodeModule)
A.CodePane.Show
End Sub

Sub JmpCls(ClsNm$)
End Sub


Sub JmpCmp(CmpNm$)
ClsWinzExlCmpOoImm CmpNm
ShwCmp CmpNm
TileV
End Sub


