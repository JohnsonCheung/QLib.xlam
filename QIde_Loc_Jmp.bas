Attribute VB_Name = "QIde_Loc_Jmp"
Option Explicit
Private Const CMod$ = "MIde_Loc_Jmp."
Private Const Asm$ = "QIde"
Sub MdJmpLno(A As CodeModule, Lno)
JmpMd A
JmpLno Lno
End Sub
Function HasMdn(Mdn, Optional Inf As Boolean) As Boolean
If HasItn(CPj.VBComponents, Mdn) Then HasMdn = True: Exit Function
If Inf Then InfLin CSub, FmtQQ("Md[?] not exist", Mdn)
End Function
Sub Jmp(MdnDotLno$)
JmpMdn BefOrAll(MdnDotLno, ".")
JmpLno Aft(MdnDotLno, ".")
End Sub
Function JmpMdn%(Mdn)
If Not HasMdn(Mdn, Inf:=True) Then JmpMdn = 1: Exit Function
JmpMd Md(Mdn)
TileV
End Function
Sub JmpMdPos(A As MdPos)
JmpMd A.Md
JmpLinPos A.LinPos
End Sub

Sub JmpLin(Lno)
With CurCdPne
    .TopLine = Lno
    .SetSelection Lno, 1, Lno, Len(CMd.Lines(Lno, 1))
End With

End Sub

Sub JmpRRCC(A As RRCC)
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

Sub JmpLno(Lno)
JmpLinPos LinPos(Lno, EmpPos)
End Sub
Function MdPoszML(A As CodeModule, Lno) As MdPos
MdPoszML = MdPos(A, LinPoszL(Lno))
End Function

Sub JmpMthzMN(A As CodeModule, Mthn)
JmpMd A
JmpMth Mthn
End Sub

Function MdPoseszMM(A As CodeModule, Mthn) As MdPoses
Dim I, IMthLin$, ILno, IPos As Pos
For Each I In MthIxyzMN(A, Mthn)
    ILno = I + 1
    IMthLin = ContLinzML(A, ILno)
    IPos = PoszSS(IMthLin, Mthn)
    PushMdPos MdPoseszMM, MdPoszMLP(A, ILno, IPos)
Next
End Function
Sub PushMdPos(O As MdPoses, M As MdPos)

End Sub
Sub PushMdPoses(O As MdPoses, M As MdPoses)
Dim J&
For J = 0 To M.N - 1
    PushMdPos O, M.Ay(J)
Next
End Sub

Function MdPoseszPM(P As VBProject, Mthn) As MdPoses
Dim M, Md As CodeModule
For Each M In Itr(MdNyzPPm(P, Mthn))
    Set Md = M
    PushMdPoses MdPoseszPM, MdPoseszMM(Md, Mthn)
Next
End Function

Sub JmpMthzP(Pj As VBProject, Mthn)
Dim A As MdPoses: A = MdPoseszPM(Pj, Mthn)
Select Case A.N
Case 0: Debug.Print FmtQQ("Mth[?] not found", Mthn)
Case 1: JmpMdPos A.Ay(0)
Case 2:
    Dim J%
    For J = 0 To A.N - 1
        Debug.Print "JmpMdPos " & MdPosStr(A.Ay(J))
    Next
End Select

End Sub
Sub JmpMth(Mthn)
JmpMthzP CPj, Mthn
End Sub

Sub JmpM()
JmpzM CMd
End Sub

Sub JmpPj(P As VBProject)
ClsWin
Dim Md As CodeModule
Set Md = FstMd(P)
If IsNothing(Md) Then
    Exit Sub
End If
Md.CodePane.Show
TileVBtn.Execute
DoEvents
End Sub

Sub JmpzMR(M As CodeModule, R As RRCC)
JmpzM M
JmpzRRCC R
End Sub

Sub JmpMd(Mdnn$)
Dim MdAy() As CodeModule: MdAy = MdAyzNN(Mdnn)
Dim I
For Each I In Itr(MdAy)
    JmpzM CvMd(I)
Next
ClsWinOfExlAp WinAyzMdAy(MdAy)
BtnOf TileV.Execute
End Sub

Sub JmpzM(A As CodeModule)
A.CodePane.Show
End Sub

Sub JmpCls(Clsn$)
End Sub

Sub JmpCmp(Cmpn$)
ClsWinOfExlCmpOoImm Cmpn
ShwCmp Cmpn
TileV
End Sub


