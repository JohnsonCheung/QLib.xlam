Attribute VB_Name = "QIde_Ens_EnsSubZ"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Ens_SubZ."
Private Const Asm$ = "QIde"


Private Function XCallgLin(Mthn, CallingPm$, PrpGetAset As Aset)
If PrpGetAset.Has(Mthn) Then
    XCallgLin = "XX = " & Mthn & "(" & CallingPm & ")"  ' The Mthn is object, no need to add [Set] XX =, the compiler will not check for this
Else
    XCallgLin = Mthn & AddPfxSpczIfNB(CallingPm)
End If
End Function

Private Function XCall(MthNy$(), ABCPm$(), Mdy$) As String()
Dim I: For Each I In Itr(IxyzSrtAy(MthNy))
    PushI XCall, MthNy(I) & " " & ABCPm(I)
Next
End Function

Private Function XCallGet(GetNy$(), ABCPm$(), GetVarNm$()) As String()
Dim I: For Each I In Itr(IxyzSrtAy(GetNy))
    PushI XCallGet, GetVarNm(I) & " = " & GetNy(I) & "(" & ABCPm(I) & ")"
Next
End Function
Private Function XCallLet(LetNy$(), ABCPm$(), LetVarNm$()) As String()
Dim I: For Each I In Itr(IxyzSrtAy(LetNy))
    PushI XCallLet, LetNy(I) & "(" & ABCPm(I) & ") = " & LetVarNm(I)
Next
End Function

Private Function XCdDim$(DclSfx$())
Dim O$()
    Dim J%: For J = 0 To UB(DclSfx)
        PushI O, "Dim A" & J & DclSfx(J)
    Next
XCdDim = JnCrLf(O)
End Function


Private Sub Z_MthSubZ()
Dim M As CodeModule ' Pm
GoSub Z
'GoSub T1
Exit Sub
Z:
    Set M = CMd
    Brw MthSubZ(M)
    Return
T1:
    Set M = Md("QIde_EnsSubZ")
    Ept = ""
    GoTo Tst
Tst:
    Act = MthSubZ(M)
    C
    Return
End Sub

Sub EnsSubZP()
EnsSubZzP CPj
EnsPrvZzP CPj
End Sub

Sub EnsSubZM()
EnsSubZ CMd
EnsPrvZ CMd
End Sub

Function MthSubZM$()
MthSubZM = MthSubZ(CMd)
End Function

Private Sub EnsSubZ(M As CodeModule)
RplMth M, "SubZ", MthSubZ(M)
End Sub

Private Sub EnsSubZzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    EnsSubZ C.CodeModule
Next
End Sub

Private Function MthSubZ$(M As CodeModule)
Dim Src$(): Src = SrczM(M)
Dim O As New Bfr
O.Var "Private Sub Z()" '<--
'-- Lbl lin-------------------------------------------------------------------------------------------------------------
O.Var Mdn(M) & ":" '<-- Som the chg <Lbl>: to <Lbl>. will have drp down
'-- ZDash---------------------------------------------------------------------------------------------------------------
'   is list of Mthn of Z_XXX within @M
'   is calling such lis of mth as the test of the M@
Dim Mth$():     Mth = MthNyzS(Src)
Dim ZDash$(): ZDash = AwPfx(Mth, "Z_")
:                     PushIAy O, SrtAy(ZDash) ' <== ZDash
:                     PushI O, "Exit Sub"     ' <== Exit SUb
'-- Var .. .. For-Cd[Dim Pub Prv] --
Dim MthD As Drs:    MthD = DoMthczS(Src)         '
Dim MthL$():        MthL = StrCol(MthD, "MthLin")
Dim PubL$()
Dim PrvL$()
Dim FrdL$()

Dim PubPm$():            PubPm = BetBktzAy(PubL)
Dim PrvPm$():            PrvPm = BetBktzAy(PrvL)
Dim FrdPm$():            FrdPm = BetBktzAy(FrdL)
Dim ArgAy$():      ArgAy = ArgAyzPmAy(PubPm)    ' Each ArgAy in ArgAy become on PubMthPm   Eg, 1-ArgAy = A$, B$, C%, D As XYZ => 4-PubMthPm
                                                         ' ArgSfxDic is Key=ArgSfx and Val=A, B, C
                                                         ' ArgSfx is ArgAy-without-Nm

Dim ArgSfx$():                          ArgSfx = SrtAy(AwDist(ArgSfxy(ArgAy)))
Dim ArgSfxToABC As Dictionary: Set ArgSfxToABC = DiKqABC(ArgSfx)

'--
Dim Pub$():          Pub = StrCol(DwEq(MthD, "Mdy", "Pub"), "Mthn")
Dim Prv$():          Prv = StrCol(DwEq(MthD, "Mdy", "Prv"), "Mthn")
Dim Frd$():          Frd = StrCol(DwEq(MthD, "Mdy", "Frd"), "Mthn")
Dim PubCPm$()
Dim FrdCPm$()
Dim PrvCPm$()
Dim GetNy$()
Dim LetNy$()
Dim GetPm$()
Dim LetPm$()
Dim GetVarNy$()
Dim LetVarNy$()
'-- Cd Dim--------------------------------------------------------------------------------------------------------------
'   Cd of pub mth so that Shf-F2 can jmp
'   Cd of prv Mth so ''
'   End Sub
O.Var XCall(Pub, PubCPm, "Pub")
O.Var XCall(Prv, PrvCPm, "Prv")
O.Var XCall(Frd, FrdCPm, "Frd")
O.Var XCallGet(GetNy, GetPm, GetVarNy)
O.Var XCallLet(LetNy, LetPm, LetVarNy)
O.Var "End Sub"
'== SubZ ===============================================================================================================
MthSubZ = O.Lines  '<==
End Function

