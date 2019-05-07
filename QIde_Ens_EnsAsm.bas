Attribute VB_Name = "QIde_Ens_EnsAsm"
Option Explicit
Private Const Asm$ = "QIde"
Private Const Ns$ = "QIde.Qualify"
Private Const CMod$ = "BEnsAsm."
Enum EmMdy
    EiIns
    EiRmv
    EiRpl
End Enum
Type MdygLinPm
    Act As EmMdy
    Lno As Long
    OldLin As String
    NewLin As String
End Type
Type SomMdygLinPm
    Som As Boolean
    Itm As MdygLinPm
End Type
Private Const A$ = 1

Private Sub Z_EnsAsmzMd()
Dim Md As CodeModule
GoSub T0
Exit Sub
T0:
    Set Md = CurMd
    GoTo Tst
Tst:
    EnsAsmzMd Md
    Return
End Sub

Sub EnsAsmM()
EnsAsmzMd CurMd
End Sub

Sub EnsAsmP()
EnsAsmzPj CurPj
End Sub
Sub EnsAsmzPj(A As VBProject)
Dim C As VBComponent, Mdyd%, Skpd%
For Each C In A.VBComponents
    If EnsAsmzMd(C.CodeModule) Then
        Mdyd = Mdyd + 1
    Else
        Skpd = Skpd + 1
    End If
Next
Inf CSub, "Done", "Pj Mdyd Skpd Tot", A.Name, Mdyd, Skpd, Mdyd + Skpd
End Sub
Function EnsAsmzMd(A As CodeModule) As Boolean 'Return True if the Module has been changed
If A.Parent.Type = vbext_ct_Document Then Exit Function
With SomMdygLinPmzSetgAsmConst(A)
    If .Som Then
        EnsAsmzMd = True
        Debug.Print MdNm(A); "<============= Mdy"
        MdyMdzLin A, .Itm
    End If
End With
End Function

Function HasAsmNm(MdNm$) As Boolean
If FstChr(MdNm) <> "M" Then Exit Function
If Not IsAscUCas(Asc(SndChr(MdNm))) Then Exit Function
HasAsmNm = True
End Function

Function AsmNm$(A As CodeModule)
Dim N$: N = MdNm(A)
If HasAsmNm(N) Then AsmNm = RplFstChr(Bef(N, "_"), "Q")
End Function

Function ConstLinzAsm$(A As CodeModule)
Dim N$: N = AsmNm(A)
If N = "" Then Exit Function
ConstLinzAsm = FmtQQ("Private Const Asm$ = ""?""", N)
End Function

Function IsEqSomMdygLinPm(A As SomMdygLinPm, B As SomMdygLinPm) As Boolean
If A.Som And B.Som And (IsEqMdygLinPm(A.Itm, B.Itm)) Then IsEqSomMdygLinPm = True
End Function

Function IsEqMdygLinPm(A As MdygLinPm, B As MdygLinPm) As Boolean
With A
Select Case True
Case .Act <> B.Act, .Lno <> B.Lno, .NewLin <> B.NewLin, .OldLin <> B.OldLin:
Case Else: IsEqMdygLinPm = True
End Select
End With
End Function

Private Sub Z_SomMdygLinPmzSetgAsmConst()
Dim Md As CodeModule, Act As SomMdygLinPm, Ept As SomMdygLinPm
GoSub T0
Exit Sub
T0:
    Set Md = CurMd
    Ept = SomInsgLinPm(2, "Private Const Asm$ = ""QIde""")
    GoTo Tst
Tst:
    Act = SomMdygLinPmzSetgAsmConst(Md)
    If Not IsEqSomMdygLinPm(Act, Ept) Then Stop
    Return
End Sub

Function RplgLinPm(Lno&, OldLin$, NewLin$) As MdygLinPm
RplgLinPm = MdygLinPm(EiRpl, Lno, OldLin, NewLin)
End Function
Function InsgLinPm(Lno&, NewLin$) As MdygLinPm
InsgLinPm = MdygLinPm(EiIns, Lno, "", NewLin)
End Function
Function RmvgLinPm(Lno&, OldLin$) As MdygLinPm
RmvLinPm = MdygLinPm(EiRmv, Lno, OldLin, "")
End Function
Function SomRplgLinPm(Lno&, OldLin$, NewLin$) As SomMdygLinPm
SomRplgLinPm = SomMdygLinPm(RplgLinPm(Lno, OldLin, NewLin))
End Function
Function SomRmvgLinPm(Lno&, OldLin$) As SomMdygLinPm
RmvInsLinPm = RmvMdygLinPm(InsgLinPm(Lno, OldLin))
End Function
Function SomInsgLinPm(Lno&, NewLin$) As SomMdygLinPm
SomInsgLinPm = SomMdygLinPm(InsgLinPm(Lno, NewLin))
End Function
Function SomMdygLinPm(Itm As MdygLinPm) As SomMdygLinPm
With SomMdygLinPm
    .Som = True
    .Itm = Itm
End With
End Function
Function MdygLinPm(Act As EmMdy, Lno&, OldLin$, NewLin$) As MdygLinPm
With MdygLinPm
    .Act = Act
    .Lno = Lno
    .OldLin = OldLin
    .NewLin = NewLin
End With
End Function
Function LnozAsmConst(A As CodeModule)
LnozAsmConst = LnozConst(A, "Asm$")
End Function
Function SomMdygLinPmzSetgAsmConst(A As CodeModule) As SomMdygLinPm
Dim NewLin$: NewLin = ConstLinzAsm(A): If NewLin = "" Then Exit Function
Dim O As SomMdygLinPm
Dim Lno&: Lno = LnozAsmConst(A)
Dim OldLin$: OldLin = ContLinzMd(A, Lno)
Select Case True
Case Lno = 0: O = SomInsgLinPm(LnozAftOptzAndImpl(A), NewLin)
Case Lno > 0 And OldLin = "": Thw CSub, "Lno>0, OldLin must have value", "Md Lno", MdNm(A), Lno
Case Lno > 0 And OldLin = NewLin:
Case Lno > 0 And OldLin <> NewLin: O = SomRplgLinPm(Lno, OldLin, NewLin)
Case Else: ThwImpossible CSub
End Select
SomMdygLinPmzSetgAsmConst = O
End Function

Function LnozConstOrAftOpt&(A As CodeModule, CnstNm$)
Dim O&: O = LnozConst(A, CnstNm)
If O > 0 Then
    LnozConstOrAftOpt = O
Else
    LnozConstOrAftOpt = LnozAftOptzAndImpl(A)
End If
End Function
Private Sub Z_LnozConst()
Dim Md As CodeModule, CnstNm$
GoSub T0
Exit Sub
T0:
    Set Md = CurMd
    CnstNm = "A$"
    Ept = 14&
    GoTo Tst
Tst:
    Act = LnozConst(Md, CnstNm)
    C
    Return
End Sub
Function LnozConst&(A As CodeModule, CnstNm$)
Dim O&, C$, L$
C = "Const " & CnstNm
For O = 1 To A.CountOfDeclarationLines
    L = RmvMdy(A.Lines(O, 1))
    If HasPfx(L, C) Then LnozConst = O: Exit Function
Next
End Function
Sub MdyMdzLin(A As CodeModule, B As MdygLinPm)
With B
Select Case .Act
Case EiIns: InsLin A, .Lno, .NewLin
Case EiRmv: RmvLin A, .Lno, .OldLin
Case EiRpl: RplLin A, .Lno, .OldLin, .NewLin
Case Else: Thw CSub, "Unexpected Act.  Should be Ins or Rpl only", "Act", Act
End Select
End With
End Sub
Sub InsLin(A As CodeModule, Lno&, Lin$)
A.InsertLines Lno, Lin
End Sub

Sub RplLin(A As CodeModule, Lno&, OldLin$, NewLin$)
If A.Lines(Lno, 1) <> OldLin Then Thw CSub, "Md-Lin <> OldLno", "Md Lno Md-Lin OldLin NewLin", MdNm(A), Lno, A.Lines(Lno, 1), OldLin
A.ReplaceLine Lno, NewLin
End Sub

Sub RmvLin(A As CodeModule, Lno&, OldLin$)
Dim N%: N = ContLinCntzMd(A, Lno)
Dim LinFmMd$: LinFmMd = A.Lines(Lno, N)
If LinFmMd <> OldLin Then Thw CSub, "Lines from Md <> OldLines", "Md Lno Lines-from-Md OldLines", MdNm(A), Lno, LinFmMd, OldLin
A.DeleteLines Lno, N
End Sub

