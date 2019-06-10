Attribute VB_Name = "QIde_Loc_Jmp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Loc_Jmp."
Private Const Asm$ = "QIde"

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

Sub Jmp(MdnDotLno$)
JmpMdn BefOrAll(MdnDotLno, ".")
JmpLin Aft(MdnDotLno, ".")
End Sub

Sub JmpMdn(Mdn)
If Not HasMdn(Mdn) Then Exit Sub
JmpzM Md(Mdn)
End Sub

Sub JmpLin(Lno)
With CPne
    .TopLine = Lno
    .SetSelection Lno, 1, Lno, Len(CMd.Lines(Lno, 1))
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

Sub JmpzPM(Pj As VBProject, Mthn)
Dim A As MdPoses: A = MdPoseszPM(Pj, Mthn)
Select Case A.N
Case 0: Debug.Print FmtQQ("Mth[?] not found", Mthn)
Case 1: 'JmpMdPos A.Ay(0)
Case 2:
    Dim J%
    For J = 0 To A.N - 1
        Debug.Print "JmpMdPos " & MdPosStr(A.Ay(J))
    Next
End Select

End Sub
Sub JmpMth(Mthn)
JmpzPM CPj, Mthn
End Sub

Sub JmpM()
JmpzM CMd
End Sub

Sub JmpzP(P As VBProject)
ClsWin
Dim Md As CodeModule
Set Md = FstMd(P)
If IsNothing(Md) Then
    Exit Sub
End If
Md.CodePane.Show
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


