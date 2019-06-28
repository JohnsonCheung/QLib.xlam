Attribute VB_Name = "QIde_Ens_EnsAsm"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const Ns$ = "QIde.Qualify"
Private Const CMod$ = "BEnsAsm."

Function AsmnzMdn$(Mdn$)
If FstChr(Mdn) = "Q" Then
    If HasSubStr(Mdn, "_") Then
        AsmnzMdn = Bef(Mdn, "_")
    End If
End If
End Function

Sub EnsAsmP()
EnsAsmzP CPj
End Sub

Function EnsAsmzM(M As CodeModule, Optional Rpt As EmRpt) As Boolean
If IsMdEmp(M) Then Exit Function
If CmpTyzM(M) = vbext_ct_Document Then Exit Function
Const T$ = "Private Const ?$ = ""?"""
'-- Fnd-CMod-Dta
Dim Mdn$:                Mdn = MdnzM(M)
Dim CModLno&:        CModLno = LnozDclCnst(M, "CMod")
Dim CModLAct$:                 If CModLno > 0 Then CModLAct = M.Lines(CModLno, 1)
Dim CModLEpt$:      CModLEpt = FmtQQ(T, "CMod", Mdn & ".")
Dim CModRpl As Drs:  CModRpl = XRpl(CModLno, CModLAct, CModLEpt)
Dim CModIns$:                  If CModLno = 0 Then CModIns = CModLEpt

'-- Fnd-Asm-Dta
Dim Asmn$:                   Asmn = AsmnzMdn(Mdn)
Dim AsmLno&:               AsmLno = LnozDclCnst(M, "Asm")
Dim AsmLAct$:                       If AsmLno > 0 Then AsmLAct = M.Lines(AsmLno, 1)
Dim AsmNoUpd As Boolean: AsmNoUpd = Asmn <> ""
If Not AsmNoUpd Then
    Dim AsmLEpt$:      AsmLEpt = FmtQQ(T, "Asm", Asmn)
    Dim AsmRpl As Drs:  AsmRpl = XRpl(AsmLno, AsmLAct, AsmLEpt)
    Dim AsmIns$:        AsmIns = AsmLEpt
End If

'-- Fnd-Ns-Dta
Dim Nsn$:                   Nsn = NsnzMdn(Mdn)
Dim NsLno&:               NsLno = LnozDclCnst(M, "Ns")
Dim NsLAct$:                      If NsLno > 0 Then NsLAct = M.Lines(NsLno, 1)
Dim NsNoUpd As Boolean: NsNoUpd = Nsn <> ""
If Not NsNoUpd Then
    Dim NsLEpt$:      NsLEpt = FmtQQ(T, "Ns", Nsn)
    Dim NsRpl As Drs:  NsRpl = XRpl(NsLno, NsLAct, NsLEpt)
    Dim NsIns$:        NsIns = NsLEpt
End If

'-- Fnd-All-Dta
Dim AllIns$():     AllIns = SyNB(CModIns, AsmIns, NsIns)
Dim AllRpl As Drs: AllRpl = AddDrs3(CModRpl, AsmRpl, NsRpl)
'== RplLin =============================================================================================================
'== InsLin =============================================================================================================
If IsUpdzRpt(Rpt) Then
    If HasReczDrs(AllRpl) Then EnsAsmzM = True Else If Si(AllIns) > 0 Then EnsAsmzM = True
    RplLin M, AllRpl
    InsLinzDcl M, AllIns
End If
Rpt:
    Dim IsRpt As Boolean, IsPush As Boolean
    IsRpt = IsRptzRpt(Rpt)
    IsPush = IsPushzRpt(Rpt)
    If IsRpt Or IsPush Then
        X "======================================"
        X Mdn
        X "Insert Const lines: Count=" & Si(AllIns)
        X TabAy(AllIns)
        X "Update Const lines: COunt=" & NReczDrs(AllRpl)
        XDrs AllRpl
        XEnd
        Dim Msg$(): Msg = XX
    End If
    If IsRpt Then Brw Msg
    If IsPush Then X Msg
End Function

Sub EnsAsmzP(P As VBProject, Optional Rpt As EmRpt)
Dim C As VBComponent, Mdyd%, Skpd%
Dim Rpt1 As EmRpt
Rpt1 = EiPushOnly
Erase XX
For Each C In P.VBComponents
    If EnsAsmzM(C.CodeModule, Rpt1) Then
    Stop
        Mdyd = Mdyd + 1
    Else
        Skpd = Skpd + 1
    End If
'    Brw XX: Stop
Next
Brw XX
Inf CSub, "Done", "Pj Mdyd Skpd Tot", P.Name, Mdyd, Skpd, Mdyd + Skpd
End Sub
Private Function XRpl(Lno&, LAct$, LEpt$) As Drs
If Lno = 0 Then Exit Function
If LAct = LEpt Then Exit Function
XRpl = LNewO(Av(Array(Lno, LEpt, LAct)))
End Function

Sub EnsCnstzMth(M As CodeModule, Mthn$, Cnstn$, NewL$)

End Sub

Function EnsLin(M As CodeModule, L&, NewL$) As Boolean
If L = 0 Then Exit Function
If M.Lines(L, 1) = NewL Then Exit Function
If NewL = "" Then
    M.DeleteLines L, 1
Else
    M.ReplaceLine L, NewL
End If
EnsLin = True
End Function

Function HasAsmn(Mdn) As Boolean
If FstChr(Mdn) <> "M" Then Exit Function
If Not IsAscUCas(Asc(SndChr(Mdn))) Then Exit Function
HasAsmn = True
End Function

Function IxzCnst&(Src$(), Cnstn$)
Dim O&, S
For Each S In Itr(Src)
    If CnstnzL(S) = Cnstn Then
        IxzCnst = O
    End If
    O = O + 1
Next
IxzCnst = -1
End Function

Function LnozDclCnst%(M As CodeModule, Cnstn$)
Dim O%, L$
Dim C$: C = "Const " & Cnstn
For O = 1 To M.CountOfDeclarationLines
    L = RmvMdy(M.Lines(O, 1))
    If ShfPfx(L, "Const ") Then
        If TakNm(L) = Cnstn Then LnozDclCnst = O: Exit Function
    End If
Next
End Function

Function LnozFstCd&(M As CodeModule)
Stop

End Function
Function LnozFstDcl&(M As CodeModule)
Dim J&
For J = 1 To M.CountOfDeclarationLines
    Dim L$: L = Trim(M.Lines(J, 1))
    If Not HasPfxss(L, "Option Implements '") Then
        If L <> "" Then
            LnozFstDcl = True
            Exit Function
        End If
    End If
Next

End Function

Function NsnzMdn$(Mdn$)
If FstChr(Mdn) = "Q" Then
    Dim A$: A = BefOrAll(Mdn, "__")
    Dim P1%: P1 = InStr(A, "_")
    If P1 = 0 Then Exit Function
    Dim P2%: P2 = InStrRev(A, "_")
    If P1 = P2 Then Exit Function
    NsnzMdn = Mid(A, P1 + 1, P2 - P1 - 1)
End If
End Function

Private Sub Z_LnozDclConst()
Dim Md As CodeModule, Cnstn$
GoSub T0
Exit Sub
T0:
    Set Md = CMd
    Cnstn = "A$"
    Ept = 14&
    GoTo Tst
Tst:
    Act = LnozDclCnst(Md, Cnstn)
    C
    Return
End Sub

Private Sub Z_AsmnzMdn()
BrwDrs DrszMapAy(Itn(CPj.VBComponents), "AsmnzMdn NsnzMdn")
End Sub

Private Sub Z()
QIde_Ens_EnsAsm:
End Sub
