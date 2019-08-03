Attribute VB_Name = "QIde_Ens_EnsSubZ"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Ens_SubZ."
Private Const Asm$ = "QIde"
Private Type DiArgItmqVar '
    D As Dictionary
End Type

Private Function WCd$(MdMdyCall As Drs, Itm$)
Dim Cl$():           Cl = StrCol(DwEq(MdMdyCall, "Mdy", Itm), "Call") ' Cl = CallLy
Dim ClSrt$():     ClSrt = SrtAy(Cl)
Dim ClAliR$():   ClAliR = AlignRzT1(ClSrt)
Dim ClCrPfx$(): ClCrPfx = AddPfxzAy(ClAliR, vbCrLf)
Dim Cd$:             Cd = Jn(ClCrPfx)
WCd = AddNB(vbCrLf, "'-- " & Itm & " -----", Cd)
End Function

Private Function XCdDim$(ArgVar$(), DclSfx$())
Const N% = 10
Dim ArgVarGp(): ArgVarGp = GpAy(ArgVar, N)
Dim DclSfxGp(): DclSfxGp = GpAy(DclSfx, N)
Dim J%, O$(): For J = 0 To UB(ArgVarGp)
    PushI O, XCdDim__OneLin(CvSy(ArgVarGp(J)), CvSy(DclSfxGp(J)))
Next
XCdDim = AddNB(vbCrLf, "'-- Dim -----", Jn(AddPfxzAy(O, vbCrLf)))
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
    Set M = Md("QIde_Ens_EnsSubZ")
    Ept = ""
    GoTo Tst
Tst:
    Act = CdSubZ(M)
    C
    Return
End Sub

Sub EnsSubZP()
EnsSubZzP CPj
End Sub

Sub EnsSubZM()
EnsSubZ CMd
End Sub

Function CdSubZM$()
CdSubZM = CdSubZ(CMd)
End Function

Private Sub EnsSubZ(M As CodeModule)
RplMth M, "Z", CdSubZ(M)
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

Private Function WCallPm$(MthPm$, D As DiArgItmqVar)
Dim O$(), Dic As Dictionary
Set Dic = D.D
Dim Arg: For Each Arg In Itr(SplitCommaSpc(MthPm))
    If Not Dic.Exists(Arg) Then Stop
    PushI O, Dic(Arg)
Next
WCallPm = JnCommaSpc(O)
End Function

Private Function WMthNy(DoMth As Drs, Mdy$) As String()
WMthNy = StrCol(DwEq(DoMth, "Mdy", Mdy), "Mthn")
End Function

Private Function XDiArgNmqArgVar(Do2 As Drs) As Dictionary
'Fm Do2 : Do<Ty-MthPm-PmLet-RHS> @@

'Insp "QIde_Ens_EnsSubZ.XDiArgNmqArgVar", "Inspect", "Oup(XDiArgNmqArgVar) Do2", XDiArgNmqArgVar, FmtDrs(Do2): Stop
End Function

Function RmvLasTermzCommaSpc$(CommaSpc$)
RmvLasTermzCommaSpc = JnCommaSpc(RmvLasEle(SplitCommaSpc(CommaSpc)))
End Function

Private Function CdSubZ$(M As CodeModule)
'Assume: No Set|Let-Prp, All Get-Prp is Pure
':VarPfx-Md: :VarPfx #Mn-Drs#
Dim Md       As Drs:       Md = AddColzMthPm(DoMth(M))                               ' :Drs<L Mdy Ty Mthn MthLin>
Dim MdNoZ    As Drs:    MdNoZ = DwNe(Md, "Mthn", "Z")
Dim MdAddRet As Drs: MdAddRet = AddColzIsRetObj(AddColzRetAs(MdNoZ))
Dim MdRetObj As Drs: MdRetObj = SelDrs(MdAddRet, "Mdy Ty Mthn MthPm RetAs IsRetObj") ' :Drs<Mdy Ty Mthn MthPm RetAs IsRetObj>

'-- CdZDashXX(Md)-------------------------------------------------------------------------------------------------------
'   is list of Mthn of Z_XXX within @M
'   is calling such lis of mth as the test of the M@
Dim Mn$():              Mn = StrCol(MdRetObj, "Mthn")
Dim MnZDash$():    MnZDash = AwPfx(Mn, "Z_")
Dim MnZDash1$():             If Si(MnZDash) > 0 Then MnZDash1 = Sy(SrtAy(MnZDash), "Exit Sub")
Dim CdZDashXX$:  CdZDashXX = JnCrLf(MnZDash1)

Dim DoArg As Drs: DoArg = XDoArg(MdRetObj) ' :Drs<ArgItm ArgNm DclSfx ArgVar DclSfx> <ArgNm DclSfx> are unique

'-- CdDim(DoArg)  ------------------------------------------------------------------------------------------------------
'Dim Di2 As Dictionary: Set Di2 = DiczDrsCC(DoArg, "ArgVar DclSfx")
Dim ArgVar$(): ArgVar = StrCol(DoArg, "ArgVar")
Dim DclSfx$(): DclSfx = StrCol(DoArg, "DclSfx")
Dim CdDim$:     CdDim = XCdDim(ArgVar, DclSfx)

'-- CdPub(MdRetObj,DoArg) ----------------------------------------------------------------------------------------------
'   CdPrv
'   CdFrd
Dim Pub$():                              Pub = WMthNy(MdRetObj, "Pub")
Dim Prv$():                              Prv = WMthNy(MdRetObj, "Prv")
Dim Frd$():                              Frd = WMthNy(MdRetObj, "Frd")
Dim D           As DiArgItmqVar:     Set D.D = DiczDrsCC(DoArg, "ArgItm ArgVar")
Dim MdAddCallPm As Drs:          MdAddCallPm = XMdAddCallPm(MdRetObj, D)                            ' :Drs<Mdy Ty Mthn MthPm RetAs IsRetObj CallPm>
Dim MdAddCall   As Drs:            MdAddCall = XMdAddCall(MdAddCallPm)                              ' :Drs<Mdy Ty Mthn MthPm RetAs IsRetObj CallPm Call>
Dim MdGet    As Drs:             MdGet = DwEqSel(MdAddCall, "Ty", "Get", "Call")
Dim MdMdyCall   As Drs:            MdMdyCall = DwNeSel(MdAddCall, "Ty", "Get", "Mdy Call")
Dim CdPub$:                            CdPub = WCd(MdMdyCall, "Pub")
Dim CdPrv$:                            CdPrv = WCd(MdMdyCall, "Prv")
Dim CdFrd$:                            CdFrd = WCd(MdMdyCall, "Frd")
Dim CdGet$:                            CdGet = AddNB("Call", vbCrLf, StrColLines(MdGet, "Call"))

'-- Cd -----------------------------------------------------------------------------------------------------------------
Dim O$()
PushI O, "Private Sub Z()"
PushI O, Mdn(M) & ":"    '<-- Som the chg <Lbl>: to <Lbl>. will have drp down
PushNB O, CdZDashXX
PushNB O, CdDim
PushNB O, CdPub
PushNB O, CdPrv
PushNB O, CdFrd
PushNB O, CdGet
PushI O, "End Sub"
CdSubZ = JnCrLf(O)
End Function

Private Function XMdAddCallPm(MdRetObj As Drs, D As DiArgItmqVar) As Drs
'Fm MdRetObj : :Drs<Mdy Ty Mthn MthPm RetAs IsRetObj>
'Ret         : :Drs<Mdy Ty Mthn MthPm RetAs IsRetObj CallPm> @@
Dim Ix%: Ix = IxzAy(MdRetObj.Fny, "MthPm", EmThw.EiThwEr)
Dim ODy()
Dim Dr: For Each Dr In Itr(MdRetObj.Dy)
    Dim MthPm$: MthPm = Dr(Ix)
    Dim CallPm$: CallPm = WCallPm(MthPm, D)
    PushI Dr, CallPm
    PushI ODy, Dr
Next
XMdAddCallPm = AddColzFFDy(MdRetObj, "CallPm", ODy)
'Insp "QIde_Ens_EnsSubZ.XMdAddCallPm", "Inspect", "Oup(XMdAddCallPm) MdRetObj D", FmtDrs(XMdAddCallPm), FmtDrs(MdRetObj), "NoFmtr(DiArgItmqVar)": Stop
End Function

Private Sub XDoArg__SetArgVar(ODy())
'ODy : :Dy<ArgItm ArgNm DclSfx Empty>
'Ret : :Dy<ArgItm ArgNm DclSfx ArgVar>  (ArgVar,DiDupArgNmqNxtIx) = \ArgNm SngArgNy DiDupArgNmqNxtIx
Dim D As New Dictionary
Dim SngArgNy$(): SngArgNy = AwSingleEle(StrColzDy(ODy, 1))
Dim J%: For J = 0 To UB(ODy)
    Dim Dr(): Dr = ODy(J)
    Dim ArgNm$: ArgNm = Dr(1)
    Dim ArgVar$: ArgVar = XDoArg__ArgVar(ArgNm, SngArgNy, D)
    Dr(3) = ArgVar
    ODy(J) = Dr
Next
End Sub

Private Function XDoArg__ArgVar$(ArgNm$, SngArgNy$(), ODiDupArgNmqNxtIx As Dictionary)
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
'Ret       : :Drs<ArgItm ArgNm DclSfx ArgVar> <ArgNm DclSfx> are unique @@
Dim Arg, ODy(): For Each Arg In Itr(ArgAy)
    Dim Itm$: Itm = ArgItm(Arg)
    Dim ArgNm$: ArgNm = Nm(Itm)
    Dim DclSfx$: DclSfx = RmvPfx(Itm, ArgNm)
    PushI ODy, Array(Itm, ArgNm, DclSfx, Empty)
Next
XDoArg__SetArgVar ODy
XDoArg__FmArgAy = DrszFF("ArgItm ArgNm DclSfx ArgVar", ODy)
'Insp "QIde_Ens_EnsSubZ.XDoArg", "Inspect", "Oup(XDoArg__FmArgAy) ArgAy", FmtDrs(XDoArg__FmArgAy), ArgAy: Stop
End Function

Private Function XDoArg(MdRetObj As Drs) As Drs
'Fm MdRetObj : :Drs<Mdy Ty Mthn MthPm RetAs IsRetObj>
'Ret         : :Drs<ArgItm ArgNm DclSfx ArgVar DclSfx> <ArgNm DclSfx> are unique @@
Dim MthPmAy$(): MthPmAy = AeBlnk(StrCol(MdRetObj, "MthPm"))
Dim ArgAy$(): ArgAy = AwDist(AyTrim(SplitComma(JnComma(MthPmAy))))
XDoArg = XDoArg__FmArgAy(ArgAy)
'Insp "QIde_Ens_EnsSubZ.XDoArg", "Inspect", "Oup(XDoArg) MdRetObj", FmtDrs(XDoArg), FmtDrs(MdRetObj): Stop
End Function

Private Function XMdAddCall__CallStmt$(Mthn$, CallPm$, Ty$, IsRetObj As Boolean)
If Ty = "Get" Then
    XMdAddCall__CallStmt = IIf(IsRetObj, "Set ", "") & "X_" & Mthn & " = " & Mthn
Else
    XMdAddCall__CallStmt = Mthn & " " & CallPm
End If
End Function

Private Function XMdAddCall__Dr(Dr, ITy%, IMthn%, ICallPm%, IIsRetObj%) As Variant()
Dim CallPm$:               CallPm = Dr(ICallPm)
Dim Mthn$:                   Mthn = Dr(IMthn)
Dim Ty$:                       Ty = Dr(ITy)
Dim IsRetObj As Boolean: IsRetObj = Dr(IIsRetObj)
Dim CallStmt$:           CallStmt = XMdAddCall__CallStmt(Mthn, CallPm, Ty, IsRetObj)
                      XMdAddCall__Dr = Av(Dr, CallStmt)
End Function

Private Function XMdAddCall(MdAddCallPm As Drs) As Drs
'Fm MdAddCallPm : :Drs<Mdy Ty Mthn MthPm RetAs IsRetObj CallPm>
'Ret            : :Drs<Mdy Ty Mthn MthPm RetAs IsRetObj CallPm Call> @@
Dim IIsRetObj%, ITy%, IMthn%, ICallPm%: AsgIx MdAddCallPm, "Ty Mthn CallPm IsRetObj", ITy, IMthn, ICallPm, IIsRetObj
Dim ODy(), Dr: For Each Dr In Itr(MdAddCallPm.Dy)
    PushI ODy, XMdAddCall__Dr(Dr, ITy, IMthn, ICallPm, IIsRetObj)
Next
XMdAddCall = AddColzFFDy(MdAddCallPm, "Call", ODy)
'Insp "QIde_Ens_EnsSubZ.XMdAddCall", "Inspect", "Oup(XMdAddCall) MdAddCallPm", FmtDrs(XMdAddCall), FmtDrs(MdAddCallPm): Stop
End Function

Private Sub Z()
QIde_Ens_EnsSubZ:
Z_CdSubZ
Exit Sub

'-- Dim -----
Dim MdMdyCall As Drs, Itm$, ArgVar$(), DclSfx$(), DclSfxGp$(), M As CodeModule, P As VBProject, DoMth As Drs, Ty$, MthPm$
Dim D As DiArgItmqVar, Mdy$, Do2 As Drs, CommaSpc$, Dr, Do1 As Drs, MdRetObj As Drs, ODy(), ArgNm$, SngArgNy$()
Dim ODiDupArgNmqNxtIx As Dictionary, ArgAy$(), Mthn$, CallPm$, IsRetObj As Boolean, ITy%, IMthn%, ICallPm%, IIsRetObj%, MdAddCallPm As Drs

'-- Pub -----
            CdSubZM
           EnsSubZM
           EnsSubZP
RmvLasTermzCommaSpc CommaSpc
           Z_CdSubZ

'-- Prv -----
           CdSubZ M
          EnsSubZ M
        EnsSubZzP P
          WCallPm MthPm, D
              WCd MdMdyCall, Itm
           WMthNy DoMth, Mdy
        WMthNyzTy DoMth, Ty
           XCdDim ArgVar, DclSfx
   XCdDim__OneLin ArgVar, DclSfxGp
  XDiArgNmqArgVar Do2
           XDoArg MdRetObj
   XDoArg__ArgVar ArgNm, SngArgNy, ODiDupArgNmqNxtIx
  XDoArg__FmArgAy ArgAy
XDoArg__SetArgVar ODy
          XMdAddCallPm MdRetObj, D
          XMdAddCall MdAddCallPm
XMdAddCall__CallStmt Mthn, CallPm, Ty, IsRetObj
      XMdAddCall__Dr Dr, ITy, IMthn, ICallPm, IIsRetObj
                Z
End Sub
