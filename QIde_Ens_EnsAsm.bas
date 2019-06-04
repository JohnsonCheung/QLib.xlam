Attribute VB_Name = "QIde_Ens_EnsAsm"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const NS$ = "QIde.Qualify"
Private Const CMod$ = "BEnsAsm."

Function MdygszON(SOld As SomLnx, SNew As SomLnx) As Mdygs
'If ShouldDltLin(SOld, SNew) Then PUshMdyLin MdygszON, MdygOfDltzSomLnx(SOld)
'If ShouldInsLin(SOld, SNew) Then PUshMdyLin MdygszON, MdygOfIns zSomLnx(NOld)
End Function

Private Function ShouldDltLin(SOld As SomLnx, SNew As SomLnx) As Boolean

End Function

Private Function ShouldInsLin(SOld As SomLnx, SNew As SomLnx) As Boolean

End Function

Function RmvNop(A As Mdygs) As Mdygs
Dim J&
For J = 0 To A.N - 1
    If A.Ay(J).Act <> EiNop Then PushMdyg RmvNop, A.Ay(J)
Next
End Function
Function Insg(Lno, Lines$) As Insg
With Insg
End With
End Function
Private Sub Z_EnsAsmzMd()
Dim Md As CodeModule
GoSub T0
Exit Sub
T0:
    Set Md = CMd
    GoTo Tst
Tst:
    EnsAsmzMd Md
    Return
End Sub

Sub EnsAsmM()
EnsAsmzMd CMd
End Sub

Sub EnsAsmP()
EnsAsmzP CPj
End Sub
Sub EnsAsmzP(P As VBProject)
Dim C As VBComponent, Mdyd%, Skpd%
For Each C In P.VBComponents
    If EnsAsmzMd(C.CodeModule) Then
        Mdyd = Mdyd + 1
    Else
        Skpd = Skpd + 1
    End If
Next
Inf CSub, "Done", "Pj Mdyd Skpd Tot", P.Name, Mdyd, Skpd, Mdyd + Skpd
End Sub
Function EnsAsmzMd(M As CodeModule) As Boolean 'Return True if the Module has been changed
If M.Parent.Type = vbext_ct_Document Then Exit Function
With MdygOfSetgAsm(M)
    If .Act <> EiNop Then
        EnsAsmzMd = True
        Debug.Print Mdn(M); "<============= Mdy"
        'MdyMdzMM  A, .Itm
    End If
End With
End Function

Function HasAsmn(Mdn) As Boolean
If FstChr(Mdn) <> "M" Then Exit Function
If Not IsAscUCas(Asc(SndChr(Mdn))) Then Exit Function
HasAsmn = True
End Function

Function Asmn$(M As CodeModule)
Dim N$: N = Mdn(M)
If HasAsmn(N) Then Asmn = RplFstChr(Bef(N, "_"), "Q")
End Function

Function CnstLinOfAsm$(M As CodeModule)
Dim N$: N = Asmn(M)
If N = "" Then Exit Function
CnstLinOfAsm = FmtQQ("Private Const Asm$ = ""?""", N)
End Function

Function IsEqMdyg(A As Mdyg, B As Mdyg) As Boolean
If A.Act <> B.Act Then Exit Function
Stop '
End Function
Function MdygOfInszLnx(A As Lnx) As Mdyg
With A
    MdygOfInszLnx = MdygOfIns(.Ix + 1, .Lin)
End With
End Function
Function MdygOfDltzLnx(A As Lnx) As Mdyg
With A
    'MdygOfInszLnx = MdygOfDlt(.Ix + 1, .Lin)
End With
End Function
Function MdygOfIns(Lno, Lines$) As Mdyg
MdygOfIns.Act = EiIns
MdygOfIns.Ins = Insg(Lno, Lines)
End Function
Function MdygOfDlt(Lno, OldLines$) As Mdyg
MdygOfDlt.Act = EiDlt
MdygOfDlt.Dlt = Dltg(Lno, OldLines)
End Function
Function Dltg(Lno, OldLines$) As Dltg
Dltg.Lno = Lno
'Dltg.OldLines = OldLines
End Function

Function LnoOfAsmCnst(M As CodeModule)
LnoOfAsmCnst = LnoOfCnstOfAftOpt(M, "Asm$")
End Function

Function MdygOfSetgAsm(M As CodeModule) As Mdyg
Dim NewLines$: NewLines = CnstLinOfAsm(M): If NewLines = "" Then Exit Function
Dim O As Mdyg
Dim Lno: Lno = LnoOfAsmCnst(M)
Dim OldLines$: OldLines = ContLinzML(M, Lno)
Select Case True
'Case Lno = 0: O = MdygOfIns(LnoOfAftOptAndImpl(A), NewLines)
Case Lno > 0 And OldLines = "": Thw CSub, "Lno>0, OldLin must have value", "Md Lno", Mdn(M), Lno
Case Lno > 0 And OldLines = NewLines:
'Case Lno > 0 And OldLines <> NewLines: O = MdygOfRpl(Lno, OldLines, NewLines)
Case Else: ThwImpossible CSub
End Select
MdygOfSetgAsm = O
End Function

Function LnoOfCnstOrAftOpt&(M As CodeModule, Cnstn$)
Dim O&: O = LnoOfCnstOfAftOpt(M, Cnstn)
If O > 0 Then
    LnoOfCnstOrAftOpt = O
Else
'    LnoOfCnstOrAftOpt = LnoOfAftOptAndImpl(M)
End If
End Function

Private Sub Z_LnoOfConst()
Dim Md As CodeModule, Cnstn$
GoSub T0
Exit Sub
T0:
    Set Md = CMd
    Cnstn = "A$"
    Ept = 14&
    GoTo Tst
Tst:
    Act = LnoOfCnstOfAftOpt(Md, Cnstn)
    C
    Return
End Sub
Function LnoOfCnstOfAftOpt&(M As CodeModule, Cnstn$)
Dim O&, C$, L$
C = "Const " & Cnstn
For O = 1 To M.CountOfDeclarationLines
    L = RmvMdy(M.Lines(O, 1))
    If HasPfx(L, C) Then LnoOfCnstOfAftOpt = O: Exit Function
Next
End Function
