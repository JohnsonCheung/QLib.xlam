Attribute VB_Name = "MxEnsSubZ"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEnsSubZ."
Private Type DiArgItmqVar '
    D As Dictionary
End Type

Private Function WCd$(MDMdyCall As Drs, Itm$)
Dim Cl$():           Cl = StrCol(DwEq(MDMdyCall, "Mdy", Itm), "Call") ' Cl = CallLy
Dim ClSrt$():     ClSrt = AySrt(Cl)
Dim ClAliR$():   ClAliR = AlignRzT1(ClSrt)
Dim ClCrPfx$(): ClCrPfx = AddPfxzAy(ClAliR, vbCrLf)
Dim Cd$:             Cd = Jn(ClCrPfx)
WCd = AddNB(vbCrLf, "'-- " & Itm & " -----", Cd)
End Function

Private Function XCDDim$(AGVar$(), DclSfx$())
Const N% = 10
Dim ArgVarGp(): ArgVarGp = GpAy(AGVar, N)
Dim DclSfxGp(): DclSfxGp = GpAy(DclSfx, N)
Dim J%, O$(): For J = 0 To UB(ArgVarGp)
    PushI O, XCDDim__OneLin(CvSy(ArgVarGp(J)), CvSy(DclSfxGp(J)))
Next
XCDDim = AddNB(vbCrLf, "'-- Dim -----", Jn(AddPfxzAy(O, vbCrLf)))
'Insp "QIde_Ens_EnsSubZ.XCDDim", "Inspect", "Oup(XCDDim) ArgVar DclSfx", XCDDim, ArgVar, DclSfx: Stop
End Function

Private Function XCDDim__OneLin$(AGVar$(), DclSfxGp$())
Dim O$()
Dim J%: For J = 0 To UB(AGVar)
    PushI O, AGVar(J) & DclSfxGp(J)
Next
XCDDim__OneLin = "Dim " & JnCommaSpc(O)
End Function

Private Sub Z_CdSubZ()
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
    Dim AItm$: AItm = ArgItm(Arg)
    If Not Dic.Exists(AItm) Then Stop
    PushI O, Dic(AItm)
Next
WCallPm = JnCommaSpc(O)
End Function

Private Function WMthNy(DoMth As Drs, Mdy$) As String()
WMthNy = StrCol(DwEq(DoMth, "Mdy", Mdy), "Mthn")
End Function

Private Function XDiArgNmqArgVar(MD2 As Drs) As Dictionary
'Fm MD2 : Drs-Ty-MthPm-PmLet-RHS> @@

'Insp "QIde_Ens_EnsSubZ.XDiArgNmqArgVar", "Inspect", "Oup(XDiArgNmqArgVar) MD2", XDiArgNmqArgVar, FmtCellDrs(MD2): Stop
End Function


Private Function CdSubZ$(M As CodeModule)
'Assume: No Set|Let-Prp, All Get-Prp is Pure
':VarPfx-Md: :VarPfx #Mth-Drs#
'':CD: :Code
'':MD: :Mth-Drs:
'':MN: :Mthn:
'':AG: :Arg:
Dim Md       As Drs:       Md = AddMthColMthPm(DoMthzM(M))                                                        ' :Drs-L Mdy Ty Mthn MthLin>
Dim MdRetObj As Drs: MdRetObj = SelDrs(AddMthColIsRetObj(AddColzRetAs(Md)), "Mdy Ty Mthn MthPm RetSfx IsRetObj") ' :Drs-Mdy Ty Mthn MthPm RetSfx IsRetObj>

'-- CdZDashXX-----------------------------------------------------------------------------------------------------------
'   is list of Mthn of Z_XXX within @M
'   is calling such lis of mth as the test of the M@
Dim MN$():             MN = StrCol(Md, "Mthn")
Dim MNZdash$():   MNZdash = AwPfx(MN, "Z_")
Dim CDExit$():     CDExit = Sy(AySrt(MNZdash), "Exit Sub")
Dim CDZdashXX$: CDZdashXX = JnCrLf(CDExit)

Dim DoArg As Drs: DoArg = XDoArg(MdRetObj) ' :Drs-ArgItm-ArgNm-DclSfx-ArgVar-DclSfx <ArgNm DclSfx> are unique

'-- CdDim(DoArg)  ------------------------------------------------------------------------------------------------------
'Dim Di2 As Dictionary: Set Di2 = DiczDrsCC(DoArg, "ArgVar DclSfx")
Dim AGVar$(): AGVar = StrCol(DoArg, "ArgVar")
Dim AGSfx$(): AGSfx = StrCol(DoArg, "DclSfx")
Dim CDDim$:     CDDim = XCDDim(AGVar, AGSfx)

'-- CdPub(MdRetObj,DoArg) ----------------------------------------------------------------------------------------------
'   CdPrv
'   CdFrd
Dim AGDi     As DiArgItmqVar: Set AGDi.D = DiczDrsCC(DoArg, "ArgItm ArgVar")
Dim MDNoZ     As Drs:              MDNoZ = DwNe(MdRetObj, "Mthn", "Z")
Dim MDCallPm  As Drs:           MDCallPm = XMDCallPm(MDNoZ, AGDi)                               ' :Drs-Mdy Ty Mthn MthPm RetSfx IsRetObj CallPm>
Dim MDCall    As Drs:             MDCall = XMDCall(MDCallPm)                                 ' :Drs-Mdy Ty Mthn MthPm RetSfx IsRetObj CallPm Call>
Dim MDGet     As Drs:              MDGet = DwEqSel(MDCall, "Ty", "Get", "Call")
Dim MDMdyCall As Drs:          MDMdyCall = DwNeSel(MDCall, "Ty", "Get", "Mdy Call")          ' :Drs-Mdy Call>
Dim CDPub$:                        CDPub = WCd(MDMdyCall, "Pub")
Dim CDPrv$:                        CDPrv = WCd(MDMdyCall, "Prv")
Dim CDFrd$:                        CDFrd = WCd(MDMdyCall, "Frd")
Dim CDGet$:                        CDGet = AddNB("Call", vbCrLf, StrColLines(MDGet, "Call"))

'':CD: :Cd -----------------------------------------------------------------------------------------------------------------
Dim O$()
PushI O, "Private Sub Z()"
PushI O, Mdn(M) & ":"    '<-- Som the chg <Lbl>: to <Lbl>. will have drp down
PushNB O, CDZdashXX
PushNB O, CDDim
PushNB O, CDPub
PushNB O, CDPrv
PushNB O, CDFrd
PushNB O, CDGet
PushI O, "End Sub"
CdSubZ = JnCrLf(O)
End Function

Private Function XMDCallPm(MDNoZ As Drs, D As DiArgItmqVar) As Drs
'Ret : :Drs-Mdy Ty Mthn MthPm RetSfx IsRetObj CallPm> @@
Dim Ix%: Ix = IxzAy(MDNoZ.Fny, "MthPm", EmThw.EiThwEr)
Dim ODy()
Dim Dr: For Each Dr In Itr(MDNoZ.Dy)
    Dim MthPm$: MthPm = Dr(Ix)
    Dim CallPm$: CallPm = WCallPm(MthPm, D)
    PushI Dr, CallPm
    PushI ODy, Dr
Next
XMDCallPm = AddColzFFDy(MDNoZ, "CallPm", ODy)
'Insp "QIde_Ens_EnsSubZ.XMDCallPm", "Inspect", "Oup(XMDCallPm) MdNoZ D", FmtCellDrs(XMDCallPm), FmtCellDrs(MdNoZ), "NoFmtr(DiArgItmqVar)": Stop
End Function

Private Sub XDoArg__SetArgVar(ODy())
'ODy : :Dy<ArgItm ArgNm DclSfx Empty>
'Ret : :Dy<ArgItm ArgNm DclSfx ArgVar>  (ArgVar,DiDupArgNmqNxtIx) = \ArgNm SngArgNy DiDupArgNmqNxtIx
Dim D As New Dictionary
Dim SngArgNy$(): SngArgNy = AwSingleEle(StrColzDy(ODy, 1))
Dim J%: For J = 0 To UB(ODy)
    Dim Dr(): Dr = ODy(J)
    Dim ArgNm$: ArgNm = Dr(1)
    Dim AGVar$: AGVar = XDoArg__ArgVar(ArgNm, SngArgNy, D)
    Dr(3) = AGVar
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
'Ret       : :Drs-ArgItm ArgNm DclSfx ArgVar> <ArgNm DclSfx> are unique @@
Dim Arg, ODy(): For Each Arg In Itr(ArgAy)
    Dim Itm$: Itm = ArgItm(Arg)
    Dim ArgNm$: ArgNm = NM(Itm)
    Dim DclSfx$: DclSfx = RmvPfx(Itm, ArgNm)
    PushI ODy, Array(Itm, ArgNm, DclSfx, Empty)
Next
XDoArg__SetArgVar ODy
XDoArg__FmArgAy = DrszFF("ArgItm ArgNm DclSfx ArgVar", ODy)
'Insp "QIde_Ens_EnsSubZ.XDoArg", "Inspect", "Oup(XDoArg__FmArgAy) ArgAy", FmtCellDrs(XDoArg__FmArgAy), ArgAy: Stop
End Function

Private Function XDoArg(MdRetObj As Drs) As Drs
'Fm MdRetObj : :Drs-Mdy Ty Mthn MthPm RetSfx IsRetObj>
'Ret         : :Drs-ArgItm ArgNm DclSfx ArgVar DclSfx> <ArgNm DclSfx> are unique @@
Dim MthPmAy$(): MthPmAy = AeBlnk(StrCol(MdRetObj, "MthPm"))
Dim ArgAy$(): ArgAy = AwDist(AmTrim(SplitComma(JnComma(MthPmAy))))
XDoArg = XDoArg__FmArgAy(ArgAy)
'Insp "QIde_Ens_EnsSubZ.XDoArg", "Inspect", "Oup(XDoArg) MdRetObj", FmtCellDrs(XDoArg), FmtCellDrs(MdRetObj): Stop
End Function

Private Function XMDCall__CallStmt$(Mthn$, CallPm$, Ty$, IsRetObj As Boolean)
If Ty = "Get" Then
    XMDCall__CallStmt = IIf(IsRetObj, "Set ", "") & "X_" & Mthn & " = " & Mthn
Else
    XMDCall__CallStmt = Mthn & " " & CallPm
End If
End Function

Private Function XMDCall__Dr(Dr, ITy%, IMthn%, ICallPm%, IIsRetObj%) As Variant()
Dim CallPm$:               CallPm = Dr(ICallPm)
Dim Mthn$:                   Mthn = Dr(IMthn)
Dim Ty$:                       Ty = Dr(ITy)
Dim IsRetObj As Boolean: IsRetObj = Dr(IIsRetObj)
Dim CallStmt$:           CallStmt = XMDCall__CallStmt(Mthn, CallPm, Ty, IsRetObj)
                      XMDCall__Dr = Av(Dr, CallStmt)
End Function

Private Function XMDCall(MDCallPm As Drs) As Drs
'Fm MdCallPm : :Drs-Mdy Ty Mthn MthPm RetSfx IsRetObj CallPm>
'Ret         : :Drs-Mdy Ty Mthn MthPm RetSfx IsRetObj CallPm Call> @@
Dim IIsRetObj%, ITy%, IMthn%, ICallPm%: AsgIx MDCallPm, "Ty Mthn CallPm IsRetObj", ITy, IMthn, ICallPm, IIsRetObj
Dim ODy(), Dr: For Each Dr In Itr(MDCallPm.Dy)
    PushI ODy, XMDCall__Dr(Dr, ITy, IMthn, ICallPm, IIsRetObj)
Next
XMDCall = AddColzFFDy(MDCallPm, "Call", ODy)
'Insp "QIde_Ens_EnsSubZ.XMDCall", "Inspect", "Oup(XMDCall) MdCallPm", FmtCellDrs(XMDCall), FmtCellDrs(MdCallPm): Stop
End Function

Private Sub Z()
QIde_Ens_EnsSubZ:
Z_CdSubZ
Exit Sub

'-- Dim -----
Dim MDMdyCall As Drs, Itm$, AGVar$(), DclSfx$(), DclSfxGp$(), M As CodeModule, P As VBProject, DoMth As Drs, Ty$, MthPm$
Dim D As DiArgItmqVar, Mdy$, MD2 As Drs, Dr, Do1 As Drs, DoMth1 As Drs, ODy(), ArgNm$, SngArgNy$(), ODiDupArgNmqNxtIx As Dictionary
Dim ArgAy$(), Mthn$, CallPm$, IsRetObj As Boolean, ITy%, IMthn%, ICallPm%, IIsRetObj%, MDCallPm As Drs

'-- Pub -----
 CdSubZM
EnsSubZM
EnsSubZP

'-- Prv -----
           CdSubZ M
          EnsSubZ M
        EnsSubZzP P
          WCallPm MthPm, D
              WCd MDMdyCall, Itm
           WMthNy DoMth, Mdy
        WMthNyzTy DoMth, Ty
           XCDDim AGVar, DclSfx
   XCDDim__OneLin AGVar, DclSfxGp
  XDiArgNmqArgVar MD2
             XDoArg DoMth1
   XDoArg__ArgVar ArgNm, SngArgNy, ODiDupArgNmqNxtIx
  XDoArg__FmArgAy ArgAy
XDoArg__SetArgVar ODy
          XMDCallPm DoMth1, D
          XMDCall MDCallPm
XMDCall__CallStmt Mthn, CallPm, Ty, IsRetObj
      XMDCall__Dr Dr, ITy, IMthn, ICallPm, IIsRetObj
                Z
         Z_CdSubZ
End Sub