Attribute VB_Name = "QIde_Ens_EnsSubZ"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Ens_SubZ."
Private Const Asm$ = "QIde"
Private Type DiArgItmqVar '
    D As Dictionary
End Type

Private Function WCdCall$(MthNy$(), Pm$(), Mdy$)
Dim O$(), I: For Each I In Itr(IxyzSrtAy(MthNy))
    PushI O, MthNy(I) & " " & Pm(I)
Next
WCdCall = XCd(O, Mdy)
End Function
Private Function XCd$(CdLy$(), Itm$)
If Si(CdLy) = 0 Then Exit Function
Dim O$(): PushI O, "-- " & Itm & " --"
PushIAy O, CdLy
PushI O, ""
XCd = JnCrLf(O)
End Function

Private Function XCdDim$(ArgVar$(), DclSfx$())
Const N% = 10
Dim ArgVarGp(): ArgVarGp = GpAy(ArgVar, N)
Dim DclSfxGp(): DclSfxGp = GpAy(DclSfx, N)
Dim J%, O$(): For J = 0 To UB(ArgVarGp)
    PushI O, XCdDim__OneLin(CvSy(ArgVarGp(J)), CvSy(DclSfxGp(J)))
Next
XCdDim = JnCrLf(O)
'Insp "QIde_Ens_EnsSubZ.XCdDim", "Inspect", "Oup(XCdDim) ArgVar DclSfx", XCdDim, ArgVar, DclSfx: Stop
End Function

Private Function XCdDim__OneLin$(ArgVar$(), DclSfxGp$())
Dim O$()
Dim J%: For J = 0 To UB(ArgVar)
    PushI O, ArgVar(J) & DclSfxGp(J)
Next
XCdDim__OneLin = "Dim " & JnCommaSpc(O)
End Function

Sub Z_CdSubZ()
Dim M As CodeModule ' Pm
GoSub Z
'GoSub T1
Exit Sub
Z:
    Set M = Md("QIde_Ens_EnsSubZ")
    Debug.Print CdSubZ(M)
    Return
T1:
    Set M = Md("QIde_EnsSubZ")
    Ept = ""
    GoTo Tst
Tst:
    Act = CdSubZ(M)
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

Function CdSubZM$()
CdSubZM = CdSubZ(CMd)
End Function

Private Sub EnsSubZ(M As CodeModule)
RplMth M, "SubZ", CdSubZ(M)
End Sub

Private Sub EnsSubZzP(P As VBProject)
Dim C As VBComponent
For Each C In P.VBComponents
    EnsSubZ C.CodeModule
Next
End Sub

Private Function WMthNyzTy(DoMth As Drs, Ty$) As String()
WMthNyzTy = StrCol(DwEq(DoMth, "Ty", Ty), "Mthn")
End Function

Private Function WCallPm(DoMth As Drs, Mdy$, D As DiArgItmqVar) As String()
Dim MthPmAy$(): MthPmAy = StrCol(DwEq(DoMth, "Mdy", Mdy), "MthPm")
Dim MthPm: For Each MthPm In Itr(MthPmAy)
    PushI WCallPm, WCallPm__Lin(MthPm, D)
Next
End Function
Private Function WCallPm__Arg$(Arg, D As DiArgItmqVar)
Dim ArgNm$:
End Function
Private Function WCallPm__Lin$(MthPm, D As DiArgItmqVar)
Dim O$()
Dim Arg: For Each Arg In Itr(SplitCommaSpc(MthPm))
    PushI O, WCallPm__Arg(Arg, D)
Next
WCallPm__Lin = JnCommaSpc(O)
End Function

Private Function WMthNy(DoMth As Drs, Mdy$) As String()
WMthNy = StrCol(DwEq(DoMth, "Mdy", Mdy), "Mthn")
End Function

Private Function XDi2(Do2 As Drs) As Dictionary
'Fm Do2 : Do<Ty-MthPm-PmLet-RHS>
'Ret    : Di<ArgNm,DclSfx> @@

'Insp "QIde_Ens_EnsSubZ.XDi2", "Inspect", "Oup(XDi2) Do2", XDi2, FmtDrs(Do2): Stop
End Function

Private Function XDiArgNmqArgVar(Do2 As Drs) As Dictionary
'Fm Do2 : Do<Ty-MthPm-PmLet-RHS> @@

'Insp "QIde_Ens_EnsSubZ.XDiArgNmqArgVar", "Inspect", "Oup(XDiArgNmqArgVar) Do2", XDiArgNmqArgVar, FmtDrs(Do2): Stop
End Function

Function RmvLasTermzCommaSpc$(CommaSpc$)
RmvLasTermzCommaSpc = JnCommaSpc(RmvLasEle(SplitCommaSpc(CommaSpc)))
End Function

Private Function XDo2__Dr(Dr) As Variant()
    Dim PmLet$
    Dim Ty$: Ty = Dr(0)
    Dim PmAy$(): PmAy = SplitComma(PmLet)
    If Ty = "Let" Or Ty = "Set" Then
        'LetRHS = Pop(Dr)
    Else
        'LetRHS = ""
        'PmLet = Pm
    End If
'    PushI Dr, Pm

End Function

Private Function XDo2(Do1 As Drs) As Drs
'Fm Do1 : Do<Ty-MthPm>
'Ret    : Do<Ty-MthPm-PmLet-RHS> @@
Dim ODry(), Dr: For Each Dr In Itr(Do1.Dy)
    PushI ODry, XDo2__Dr(Dr)
Next
XDo2 = DrszFF("Ty Pm LetRHS", ODry)
'Insp "QIde_Ens_EnsSubZ.XDo2", "Inspect", "Oup(XDo2) Do1", FmtDrs(XDo2), FmtDrs(Do1): Stop
End Function

Private Function CdSubZ$(M As CodeModule)
'Assume: No Set|Let-Prp, All Get-Prp is Pure
Dim Src$():           Src = SrczM(M)
Dim DoMth  As Drs:  DoMth = AddColzMthPm(DoMthzS(Src))
Dim DoMthX As Drs: DoMthX = SelDrs(DoMth, "Mdy Ty Mthn MthPm") ' :Drs<Mdy Ty Mthn MthPm>

'-- CdZDashXX-----------------------------------------------------------------------------------------------------------
'   is list of Mthn of Z_XXX within @M
'   is calling such lis of mth as the test of the M@
Dim Mth$():             Mth = StrCol(DoMth, "Mthn")
Dim ZDash$():         ZDash = AwPfx(Mth, "Z_")
Dim A$():                 A = Sy(SrtAy(ZDash), "Exit Sub")
Dim CdZDashXX$:   CdZDashXX = JnCrLf(A)
Dim DoArg As Drs:     DoArg = XDoArg(DoMthX)               ' :Drs<ArgNm DclSfx ArgVar DclSfx> <ArgNm DclSfx> are unique

'-- CdDim --------------------------------------------------------------------------------------------------------------
'Dim Di2 As Dictionary: Set Di2 = DiczDrsCC(DoArg, "ArgVar DclSfx")
Dim ArgVar$(): ArgVar = StrCol(DoArg, "ArgVar")
Dim DclSfx$(): DclSfx = StrCol(DoArg, "DclSfx")
Dim CdDim$:     CdDim = XCdDim(ArgVar, DclSfx)

'-- CdPub --------------------------------------------------------------------------------------------------------------
'   CdPrv
'   CdFrd
Dim Pub$():                Pub = WMthNy(DoMthX, "Pub")
Dim Prv$():                Prv = WMthNy(DoMthX, "Prv")
Dim Frd$():                Frd = WMthNy(DoMthX, "Frd")
Dim Di1 As Dictionary: Set Di1 = DiczDrsCC(DoArg, "ArgNm ArgVar")
Dim PubPm$():            PubPm = WCallPm(DoMthX, "Pub", Di1)
Stop
Dim PrvPm$():            PrvPm = WCallPm(DoMthX, "Prv", Di1)
Dim FrdPm$():            FrdPm = WCallPm(DoMthX, "Frd", Di1)
Dim CdPub$:              CdPub = WCdCall(Pub, PubPm, "Pub")
Dim CdPrv$:              CdPrv = WCdCall(Prv, PrvPm, "Prv")
Dim CdFrd$:              CdFrd = WCdCall(Frd, FrdPm, "Frd")

'-- CdGet --------------------------------------------------------------------------------------------------------------
Dim GetNy$():   GetNy = WMthNyzTy(DoMthX, "Get")
Dim GetLHS$(): GetLHS = XGetLHS(DoMthX)
Dim CdGet$:     CdGet = XCdGet(GetLHS, GetNy)

'-- Cd -----------------------------------------------------------------------------------------------------------------
Dim O$()
PushI O, "Private Sub Z()"
PushI O, Mdn(M) & ":"    '<-- Som the chg <Lbl>: to <Lbl>. will have drp down
PushI O, CdZDashXX
PushI O, CdDim
PushI O, CdPub
PushI O, CdPrv
PushI O, CdFrd
PushI O, CdGet
PushI O, "End Sub"
CdSubZ = JnCrLf(O)
End Function

Private Function XCdGet$(GetLHS$(), GetNy$())
'Insp "QIde_Ens_EnsSubZ.XCdGet", "Inspect", "Oup(XCdGet) GetLHS GetNy", XCdGet, GetLHS, GetNy: Stop
End Function

Private Function XGetLHS(DoMthX As Drs) As String()
'Fm DoMthX : :Drs<Mdy Ty Mthn MthPm> @@

'Insp "QIde_Ens_EnsSubZ.XGetLHS", "Inspect", "Oup(XGetLHS) DoMthX", XGetLHS, FmtDrs(DoMthX): Stop
End Function

Private Function XDi1(Do2 As Drs) As Dictionary
'Fm Do2 : Do<Ty-MthPm-PmLet-RHS>
'Ret    : Di<ArgNm,ArgVar> @@
'Insp "QIde_Ens_EnsSubZ.XDi1", "Inspect", "Oup(XDi1) Do2", XDi1, FmtDrs(Do2): Stop
End Function

Private Sub XDoArg__SetArgVar(ODy())
'ODy : :Dy<ArgNm DclSfx Empty>
'Ret : :Dy<ArgNm DclSfx ArgVar>  (ArgVar,DiDupArgNmqNxtIx) = \ArgNm SngArgNy DiDupArgNmqNxtIx
Dim D As New Dictionary
Dim SngArgNy$(): SngArgNy = AwSingleEle(StrColzDy(ODy, 0))
Dim J%: For J = 0 To UB(ODy)
    Dim Dr(): Dr = ODy(J)
    Dim ArgNm$: ArgNm = Dr(0)
    Dim ArgVar$: ArgVar = XDoArg__ArgVar(ArgNm, SngArgNy, D)
    Dr(2) = ArgVar
    ODy(J) = Dr
Next
End Sub

Function XDoArg__ArgVar$(ArgNm$, SngArgNy$(), ODiDupArgNmqNxtIx As Dictionary)
If HasEle(SngArgNy, ArgNm) Then XDoArg__ArgVar = ArgNm: Exit Function
If ODiDupArgNmqNxtIx.Exists(ArgNm) Then
    Dim I%: I = ODiDupArgNmqNxtIx(ArgNm)
    XDoArg__ArgVar = ArgNm & "_" & I
    ODiDupArgNmqNxtIx(ArgNm) = I + 1
Else
    XDoArg__ArgVar = ArgNm & "_1"
    ODiDupArgNmqNxtIx(ArgNm) = 2
End If
End Function

Private Function XDoArg__FmArgAy(ArgAy$()) As Drs
'Fm DoMthX : :Drs<Mdy Ty Mthn MthPm>
'Ret       : :Drs<ArgNm DclSfx ArgVar> <ArgNm DclSfx> are unique @@
Dim Arg, ODy(): For Each Arg In Itr(ArgAy)
    Dim Itm$: Itm = DclItm(Arg)
    Dim ArgNm$: ArgNm = Nm(Itm)
    Dim DclSfx$: DclSfx = RmvPfx(Itm, ArgNm)
    PushI ODy, Array(ArgNm, DclSfx, Empty)
Next
XDoArg__SetArgVar ODy
XDoArg__FmArgAy = DrszFF("ArgNm DclSfx ArgVar", ODy)
'Insp "QIde_Ens_EnsSubZ.XDoArg", "Inspect", "Oup(XDoArg__FmArgAy) ArgAy", FmtDrs(XDoArg__FmArgAy), ArgAy: Stop
End Function

Private Function XDoArg(DoMthX As Drs) As Drs
'Fm DoMthX : :Drs<Mdy Ty Mthn MthPm>
'Ret       : :Drs<ArgNm DclSfx ArgVar> <ArgNm DclSfx> are unique @@
Dim MthPmAy$(): MthPmAy = AeBlnk(StrCol(DoMthX, "MthPm"))
Dim ArgAy$(): ArgAy = AwDist(AyTrim(SplitComma(JnComma(MthPmAy))))
XDoArg = XDoArg__FmArgAy(ArgAy)
End Function
