Attribute VB_Name = "QIde_Ens_SubZ"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Ens_SubZ."
Private Const Asm$ = "QIde"

Private Function CdCallPMth$(Src$())
'Ret : Cd of calling Pub-Mth with Dim lines.  So the Shf-F2 will jmp to that mth
Dim PubMthL$():        PubMthL = MthLinAyzS(Src)         '                                                                                           # Pub-Mth-Lin
Dim PubGet As Aset: Set PubGet = XPubGet(PubMthL)
Dim PubMthPm$():      PubMthPm = BetBktzAy(PubMthL)
Dim PubMthN$():        PubMthN = MthNyzMthLinAy(PubMthL)
Dim ArgAy$():            ArgAy = ArgAyzPmAy(PubMthPm)    ' Each ArgAy in ArgAy become on PubMthPm   Eg, 1-ArgAy = A$, B$, C%, D As XYZ => 4-PubMthPm
                                                         ' ArgSfxDic is Key=ArgSfx and Val=A, B, C
                                                         ' ArgSfx is ArgAy-without-Nm

Dim ArgSfx$():                          ArgSfx = SrtAy(AywDist(ArgSfxy(ArgAy)))
Dim ArgSfxToABC As Dictionary: Set ArgSfxToABC = DiczEleToABC(ArgSfx)
Dim CallgPm$():                        CallgPm = XCallgPm(PubMthPm, ArgSfxToABC)
Dim HasPrp      As Boolean:             HasPrp = PubGet.Cnt > 0
Dim DimLy$():                            DimLy = XDimLy(ArgSfxToABC, HasPrp)                      ' 1-ArgAy => 1-DimLin
Dim CallgLy$():                        CallgLy = XCallgLy(PubMthN, PubMthPm, ArgSfxToABC, PubGet)
Erase XX
    X "'== Callg pub mth =="
    X DimLy
    X CallgLy
CdCallPMth = JnCrLf(XX)
End Function

Private Function XCallgLin(Mthn, CallingPm$, PrpGetAset As Aset)
If PrpGetAset.Has(Mthn) Then
    XCallgLin = "XX = " & Mthn & "(" & CallingPm & ")"  ' The Mthn is object, no need to add [Set] XX =, the compiler will not check for this
Else
    XCallgLin = Mthn & AddPfxSpczIfNB(CallingPm)
End If
End Function

Private Function XCallgLy(MthNy$(), PmAy$(), ArgDic As Dictionary, PrpGetAset As Aset) As String()
'A$() & PmAy$() are same sz
'ArgDic: Key is ArgSfx(Arg-without-Name), Val is A,B,..
'CallingLin is {Mthn} A,B,C,...
'PrpGetAset    is PrpNm set
Dim Mthn, J%, O$(): For Each Mthn In Itr(MthNy)
    Dim Pm$:               Pm = PmAy(J)
    Dim CallingPm$: CallingPm = XCallgPmzPm(Pm, ArgDic)
    PushI O, XCallgLin(Mthn, CallingPm, PrpGetAset)
    J = J + 1
Next
XCallgLy = QSrt1(O)
End Function

Private Function XCallgPmzPm$(Pm, ArgDic As Dictionary)
Dim O$(), Arg
For Each Arg In Itr(AyTrim(SplitComma(Pm)))
    PushI O, ArgDic(ArgSfx(Arg))
Next
XCallgPmzPm = JnCommaSpc(O)
End Function

Private Function XCallgPm(PmAy$(), ArgDic As Dictionary) As String()
Dim Pm
For Each Pm In Itr(PmAy)
    PushI XCallgPm, XCallgPmzPm(Pm, ArgDic)
Next
End Function

Private Function XDimLy(ArgDic As Dictionary, HasPrp As Boolean) As String()  '1-Arg => 1-DimLin
Dim ArgSfx, S$
For Each ArgSfx In ArgDic.Keys
    If HasPfx(ArgSfx, "As ") Then
        S = " "
    Else
        S = ""
    End If
    PushI XDimLy, "Dim " & ArgDic(ArgSfx) & S & ArgSfx
Next
If HasPrp Then PushI XDimLy, "Dim XX"
End Function

Private Function XPubGet(MthDclAy$()) As Aset
Dim Lin, O As Aset
Set O = EmpAset
For Each Lin In Itr(MthDclAy)
'    If IsLinPrp(Lin) Then AsetPush O, Mthn(Lin)
Next
Set XPubGet = O
End Function

Private Sub Z_MthSubZ()
Dim S$()
GoSub Z
'GoSub T1
Exit Sub
Z:
    Brw MthSubZ(CSrc)
    Return
T1:
    S = SrczMdn("MVb_Dic")
    Ept = ""
    GoTo Tst
Tst:
    Act = MthSubZ(S)
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

Private Sub EnsSubZ(M As CodeModule)
RplMth M, "SubZ", MthSubZ(Src(M))
End Sub

Private Sub EnsSubZzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    EnsSubZ C.CodeModule
Next
End Sub
Private Function MthSubZ$(Src$())
Dim O$()
PushI O, "Private Sub Z()"
PushI O, CdCallZDash(Src)
PushI O, ""
PushI O, CdCallPMth(Src)
PushI O, "End Sub"
MthSubZ = JnCrLf(O)
End Function

Private Function CdCallZDash$(Src$())
'SubZ is [Mth-`Sub Z()`-Lines], each line is calling a Z_XX, where Z_XX is a testing function
Dim M$(): M = MthNyzS(Src)
Dim ZDash$(): ZDash = AywPfx(M, "Z_")
Dim S$(): S = SrtAy(ZDash)
PushI S, "Exit Sub"
PushI S, ""
CdCallZDash = JnCrLf(S)
End Function

