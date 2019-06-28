Attribute VB_Name = "QIde_F_Loc"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Loc."
Private Const Asm$ = "QIde"

Function DoMthPos(Mthn) As Drs
Dim A As Drs: A = DwEq(DoMthP, "Mthn", Mthn)
End Function

Sub JmpzML(M As CodeModule, Lno)
JmpzM M
JmpLin Lno
End Sub

Function HasMdnzP(P As VBProject, Mdn, Optional Inf As Boolean) As Boolean
If HasItn(CPj.VBComponents, Mdn) Then HasMdnzP = True: Exit Function
If Inf Then InfLin CSub, FmtQQ("Md[?] not exist", Mdn)
End Function

Function HasMdn(Mdn, Optional Inf As Boolean) As Boolean
HasMdn = HasMdnzP(CPj, Mdn, Inf)
End Function

Sub Z_Jmp()
Jmp "QIde_Md.10"
End Sub

Sub Jmp(MdnDotLno$)
JmpMdn BefOrAll(MdnDotLno, ".")
JmpLin Aft(MdnDotLno, ".")
End Sub

Sub JmpRCC(R&, C1%, C2%)
CPne.SetSelection R, C1, R, C2
End Sub

Sub JmpMdn(Mdn)
If Not HasMdn(Mdn) Then Debug.Print "Mdn not exist": Exit Sub
JmpzM Md(Mdn)
End Sub

Sub JmpLin(Lno)
Dim L&: L = CLng(Lno)
Dim C2%: C2 = Len(CMd.Lines(L, 1)) + 1
With CPne
    .TopLine = Lno
    .SetSelection Lno, 1, Lno, C2
End With

End Sub

Sub JmpRRCC(A As RRCC)
Dim L&, C1%, C2%
With CPne
    If C1 = 0 Or C2 = 0 Then
        C1 = 1
        C2 = Len(.CodeModule.Lines(L, 1)) + 1
    End If
    .TopLine = L
    .SetSelection L, C1, L, C2
End With
'SendKeys "^{F4}"
End Sub

Sub JmpMthzMN(M As CodeModule, Mthn)
JmpzM M
JmpMth Mthn
End Sub

Function MdPoseszMM(M As CodeModule, Mthn) As MdPoses
Dim I, IMthLin$, ILno, IPos As Pos
For Each I In MthIxyzMN(M, Mthn)
    ILno = I + 1
    IMthLin = ContLinzML(M, ILno)
    IPos = PoszSS(IMthLin, Mthn)
    PushMdPos MdPoseszMM, MdPoszMLP(M, ILno, IPos)
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

Sub JmpMth(Mthn)
Dim M As CodeModule
Dim L&: L = MthLnozMM(M, Mthn)
JmpzML M, L
End Sub

Sub JmpM()
JmpzM CMd
End Sub

Sub JmpzP(P As VBProject)
ClsWin
Dim M As CodeModule
Set M = FstMd(P)
If IsNothing(M) Then Exit Sub
JmpzM M
TileV
DoEvents
End Sub
Sub JmpzRRCC(A As RRCC)

End Sub
Sub JmpzMR(M As CodeModule, R As RRCC)
JmpzM M
JmpzRRCC R
End Sub
Function WinyzMdAy(MdAy() As CodeModule) As VBIDE.Window()

End Function
Sub JmpMdnn(Mdnn$)
Dim MdAy() As CodeModule: MdAy = MdAyzNN(Mdnn)
Dim I
For Each I In Itr(MdAy)
    JmpzM CvMd(I)
Next
ClsWinExlAp WinyzMdAy(MdAy)
TileV
End Sub

Sub JmpzM(M As CodeModule)
M.CodePane.Show
End Sub



