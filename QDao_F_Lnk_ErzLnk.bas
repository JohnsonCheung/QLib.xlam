Attribute VB_Name = "QDao_F_Lnk_ErzLnk"
Option Compare Text
Option Explicit
Private Function XIpfExi(Ipf As Drs) As Drs
'Fm  Ipf    : L Inpn Ffn IsFx HasFfn
'Ret        : L Inpn Ffn IsFx @@
XIpfExi = DwEqE(Ipf, "HasFfn", True)
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpfExi) IpfExi Ipf",FmtDrs(IpfExi), FmtDrs(IpfExi), FmtDrs(Ipf): Stop
End Function

Private Function XIpbFb(Ipf As Drs) As String()
'Fm  Ipf : L Inpn Ffn IsFx HasFfn @@
XIpbFb = StrColzEq(Ipf, "IsFx", False, "Inpn")
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbFb) IpbFb Ipf",IpbFb, IpbFb, FmtDrs(Ipf): Stop
End Function

Function ErzLnk(LnkImpSrc$()) As String()
'Fm: *InpFilSrc::SSAy{Inpn Ffn}
'Fm: *LnkImpSrc::IndentedLy{
'                  TblFx: {Fxt} [{Fxn}[.{Wsn}]] [{Stru}]
'                  TblFb: {Fbn} {Fbtt}
'                  Stru.{Stru}: {F} [{Ty}] [{Extn}]
'                  Tbl.Where: {T} {Bexp}
'                } @@
'-----------------------------------------------------------------------------------------------------------------------
Dim A$, C$
  A = 12323
  C = 234234
If A Then C = 234234
Dim D1 As Drs, D2 As Drs, D3 As Drs
Const CmPfx$ = "X"
Dim Ip  As DoLTDH:  Ip = DoLTDH(LnkImpSrc) ' L T1 Dta IsHdr
Dim Ipf As Drs:   Ipf = XIpf(Ip)         ' L Inpn Ffn IsFx HasFfn
Dim IpxHasFx As Drs:
Dim IpfFx    As Drs:    IpfFx = DwEqE(Ipf, "IsFx", True)                ' L Inpn Ffn HasFfn
Dim IpfExi   As Drs:   IpfExi = XIpfExi(Ipf)                                    ' L Inpn Ffn IsFx
Dim IpbFb$():           IpbFb = XIpbFb(Ipf)
Dim Ipw      As Drs:      Ipw = XIpw(Ip)
                           D1 = DwEqE(IpxHasFx, "IsFx", True)
                           D2 = DwEqE(D1, "HasFx", True)
                           D3 = DwEqE(D2, "HasInp", True)
Dim Ipx      As Drs:      Ipx = XIpx(Ip)
Dim IpxExiFx As Drs: IpxExiFx = D3                                              ' L T Fxn Ws Stru Fx
Dim ActWs    As Drs:    ActWs = XActWs(IpxExiFx)                                ' Fxn Ws
Dim IpxHasWs As Drs: IpxHasWs = AddColzExiB(IpxExiFx, ActWs, "Fxn Ws", "HasWs") ' T Fxn Ws Stru Fx HasFx HasWs
Dim IpxExi   As Drs:   IpxExi = DwEqE(IpxHasWs, "HasWs", True)          ' IpxExi
Dim IpxMis   As Drs:   IpxMis = DwEqE(IpxHasWs, "HasWs", False)

Dim Ips As Drs:     Ips = XIps(Ip)                ' L Stru F Ty E
Dim IpsStru$(): IpsStru = StrCol(Ips, "Stru")

Dim Ipb As Drs:    Ipb = DoLTT(Ip, "FbTbl", "Fbn Fbtt").D ' L Fbn Fbtt
Dim IpxTny$():  IpxTny = XIpxTny(Ipx)
Dim IpbTny$():  IpbTny = XIpbTny(Ipb)
Dim Tny$():        Tny = Sy(IpxTny, IpbTny)

Dim ActWsf  As Drs:  ActWsf = XActWsf(IpxExi)
Dim Ipxf    As Drs:    Ipxf = XIpxf(IpxExi, Ips)
Dim IpxfMis As Drs: IpxfMis = XIpxfMis(Ipxf, ActWsf)

'== Error===============================================================================================================
'== Er Inp (Ei)=========================================================================================================
Dim EiInpnDup$(): EiInpnDup = XEiInpnDup(Ipf)
Dim EiFfnDup$():   EiFfnDup = XEiFfnDup(Ipf)
Dim EiFfnMis$():   EiFfnMis = XEiFfnMis(Ipf)
Dim I$():                 I = Sy(EiInpnDup, EiFfnDup, EiFfnMis)

'== Er TblFx (Ex)=======================================================================================================
Dim ExTblDup$():         ExTblDup = XExTblDup(Ipx)
Dim ExFxnDup$():         ExFxnDup = XExFxnDup(Ipx)
Dim ExFxnMis$():         ExFxnMis = XExFxnMis(Ipx, IpfExi)
Dim ExWsMis$():           ExWsMis = XExWsMis(IpxMis, ActWs)
Dim ExWsMisFld$():     ExWsMisFld = XExWsMisFld(IpxfMis, ActWsf)
Dim ExWsMisFldTy$(): ExWsMisFldTy = XExWsMisFldTy(Ipxf, ActWsf)
Dim ExStruMis$():       ExStruMis = XExStruMis(Ipx, IpsStru)
Dim X$():                       X = Sy(ExFxnDup, ExFxnMis, ExWsMis, ExWsMisFld, ExWsMisFldTy)

'== Er TblFb (Eb)=======================================================================================================
Dim EbFbnDup$():   EbFbnDup = XEbFbnDup(Ipb)
Dim EbFbnMis$():   EbFbnMis = XEbFbnMis(Ipb, IpbFb)
Dim EbTblDup$():   EbTblDup = XEbTblDup(Ipb)
Dim EbTblMis$():   EbTblMis = XEbTblMis(Ipb)
Dim EbStruMis$(): EbStruMis = XEbStruMis(IpbTny, IpsStru)                           ' Use IpbTny stru to find in IpsStru
Dim B$():                 B = Sy(EbFbnDup, EbFbnMis, EbTblDup, EbTblMis, EbStruMis)

'== Er Stru (Es)========================================================================================================
Dim IpsHdStru$():
Dim IpbxStru$():
Dim EsSDup$():      EsSDup = XEsSDup(IpsHdStru)
Dim EsSMis$():      EsSMis = XEsSMis(IpsHdStru, IpbxStru)
Dim EsSExa$():      EsSExa = XEsSExa(IpsHdStru, IpbxStru)
Dim EsSNoFld$():  EsSNoFld = XEsSNoFld(IpsHdStru, IpsStru)
Dim EsFldDup$():  EsFldDup = XEsFldDup(Ips)
Dim EsTyEr$():      EsTyEr = XEsTyEr
Dim S$():                S = Sy(EsSDup, EsSMis, EsSExa, EsSNoFld, EsFldDup, EsTyEr)

'== Er TblWhere (Ew)====================================================================================================
Dim EwTblDup$():     EwTblDup = XEwTblDup(Ipw)
Dim EwTblExa$():     EwTblExa = XEwTblExa(Ipw, Tny)                ' tbl.wh is more
Dim EwBexpEmp$(): EwBexpEmp = XEwBexpEmp(Ipw)                   ' with tbl nm but no Bexp
Dim W$():                   W = Sy(EwTblDup, EwTblExa, EwBexpEmp)

'== Er Other (Eo)=======================================================================================================
Dim EoNoFxAndNoFb$: EoNoFxAndNoFb = XEoNoFxAndNoFb(Ipx, Ipb)
Dim EoHdrEr$():           EoHdrEr = XEoHdrEr(Ip)
Dim O$():                       O = Sy(EoNoFxAndNoFb, EoHdrEr)

ErzLnk = Sy(I, X, B, S, W, O)
End Function

Private Function XIpx(Ip As DoLTDH) As Drs

'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpx) Ipx Ip",FmtDrs(Ipx), FmtDrs(Ipx), FmtDrs(Ip.D): Stop
End Function

Private Function XIpfFx(Ip As DoLTDH) As Drs

End Function

Private Function XIpxf(IpxExi As Drs, Ips As Drs) As Drs
'Fm  IpxExi : IpxExi
'Fm  Ips    : L Stru F Ty E @@
Dim O As Drs, Dr, E$, IxF%, IE%, J&
O = JnDrs(IpxExi, Ips, "Stru", "F Ty E")
IxF = IxzAy(O.Fny, "F")
IE = IxzAy(O.Fny, "E")
For Each Dr In Itr(O.Dy)
    E = Dr(IE)
    If E = "" Then
        Dr(IE) = Dr(IxF)
        O.Dy(J) = Dr
    End If
    J = J + 1
Next
XIpxf = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpxf) Ipxf IpxExi Ips",FmtDrs(Ipxf), FmtDrs(Ipxf), FmtDrs(IpxExi), FmtDrs(Ips): Stop
End Function
Private Function XIpxfMisFld() As Drs

End Function
Private Function XIsShtTy(ShtTy$, IsIpsFx As Boolean) As Boolean
Select Case True
Case Not IsIpsFx And ShtTy = "": XIsShtTy = True
Case Else: XIsShtTy = IsShtTy(ShtTy)
End Select
End Function
'==================================================
Private Function XDTyEr(Ips1 As Drs) As Drs
'Fm : @Ips1::Drs{L Stru F Ty E IsIpsFx}
'Ret::*DTyEr::Drs{TyEr L Stru F E}
Dim Fny$(), Dr, ITy%, IL%, IStru%, IxF%, IE%, IxIsIpsFx%
Dim Ty$, L&, Stru$, F$, E$, IsIpsFx As Boolean, ODy()
Fny = Ips1.Fny
IL = IxzAy(Fny, "L")
IStru = IxzAy(Fny, "Stru")
IxF = IxzAy(Fny, "F")
ITy = IxzAy(Fny, "Ty")
IE = IxzAy(Fny, "E")
IxIsIpsFx = IxzAy(Fny, "IsIpsFx")
For Each Dr In Itr(Ips1.Dy)
    Ty = Dr(ITy)
    IsIpsFx = Dr(IxIsIpsFx)
    L = Dr(IL)
    Stru = Dr(IStru)
    F = Dr(IxF)
    E = Dr(IE)
    If Not XIsShtTy(Ty, IsIpsFx) Then
        PushI ODy, Array(Ty, L, Stru, F, E, IsIpsFx)
    End If
Next
XDTyEr = DrszFF("TyEr L Stru F E IsIpsFx", ODy)
End Function

Private Function WActTbl(IpfFb As Drs) As Drs
Dim Dr, T, J%, IFbn$, IFb$, Dy()
For Each Dr In Itr(IpfFb.Dy)
    IFbn = Dr(1)
    IFb = Dr(2)
    For Each T In Tni(Db(IFb))
        PushI Dy, Array(IFbn, T)
    Next
Next
WActTbl = DrszFF("Fbn T", Dy)
End Function

Private Function XActWs(IpxExiFx As Drs) As Drs
'Fm  IpxExiFx : L T Fxn Ws Stru Fx
'Ret          : Fxn Ws @@
Dim A As Drs, Dr, Fxn$, Fx$, Wsn, Dy()
'A = DwIpst(IpxExiFx, "Fxn Fx")
For Each Dr In Itr(A.Dy)
    Fxn = Dr(0)
    Fx = Dr(1)
    For Each Wsn In Wny(Fx)
        PushI Dy, Array(Fxn, Wsn)
    Next
Next
XActWs = DrszFF("Fxn Ws", Dy)
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XActWs) ActWs IpxExiFx",FmtDrs(ActWs), FmtDrs(ActWs), FmtDrs(IpxExiFx): Stop
End Function
Private Function XExWsMisFld(IpxfMis As Drs, ActWsf As Drs) As String()
If NoReczDrs(IpxfMis) Then Exit Function
Dim OFx$(), OFxn$(), OWs$(), O$(), Fxn, Fx$, Ws$, Mis As Drs, Act As Drs, J%, O1$()
AsgCol IpxfMis, "Fxn Fx Ws", OFxn, OFx, OWs
'====
PushI O, "Some columns in ws is missing"
For Each Fxn In OFxn
    Fxn = OFxn(J)
    Fx = OFx(J)
    Ws = OWs(J)
    Mis = Dw3EqE(IpxfMis, "Fxn Fx Ws", Fxn, Fx, Ws)
    Act = Dw3EqE(ActWsf, "Fxn Fx Ws", Fxn, Fx, Ws)
    '-
    
    X "Fxn    : " & Fxn
    X "Fx pth : " & Pth(Fx)
    X "Fx fn  : " & Fn(Fx)
    X "Ws     : " & Ws
    X LyzNmDrs("Mis col: ", Mis)
    X LyzNmDrs("Act col: ", Act)
    PushIAy O, TabAy(XX)
    J = J + 1
Next
XExWsMisFld = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XExWsMisFld) ExWsMisFld IpxfMis ActWsf",ExWsMisFld, ExWsMisFld, FmtDrs(IpxfMis), FmtDrs(ActWsf): Stop
End Function
Private Function XExWsMisFldTy(Ipxf As Drs, ActWsf As Drs) As String()
'Fm IpxFld : Fxn Ws Stru Ipxf Ty Fx ! Where HasFx and HasWs and Not HasFld
'Fm WsActf : Fxn Ws Ipxf Ty @@
'Dim OFxn$(), J%, Fxn$, Fx$, Act$(), Lno&(), Ws$(), ActWsf()
'OFxn = AwDist(StrCol(IpXB, "Fxn"))
''====
'If Si(OFxn) = 0 Then Exit Function
'PushI XExWsMis, "Some expected ws not found"
'For J = 0 To UB(OFxn)
'    Fxn = OFxn(J)
'    Fx = VzColEqSel(IpXB, "Fxn", Fxn, "Fx")
'    ActWsf = DwEqSel(IpXB, "Fxn", Fxn, "L Ws").Dy
'    Lno = LngAyzDyC(ActWsf, 0)
'    Ws = SyzDyC(ActWsf, 1)
'
'    Act = RmvT1zAy(AwT1(WsAct, Fxn)) '*WsActPerFxn::Sy{WsAct}
'    PushIAy XExWsMis, XMisWs_OneFx(Fxn, Fx, Lno, Ws, Act)
'Next
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XExWsMisFldTy) ExWsMisFldTy Ipxf ActWsf",ExWsMisFldTy, ExWsMisFldTy, FmtDrs(Ipxf), FmtDrs(ActWsf): Stop
End Function

Private Function XExWsMis(IpxMis As Drs, ActWs As Drs) As String()
'Fm  ActWs : Fxn Ws @@
Dim OFxn$(), J%, Fxn$, Fx$, Act$(), Lno&(), Ws$(), ActWsnn$, IpxMisi As Drs, O$()
OFxn = AwDist(StrCol(IpxMis, "Fxn"))
'====
If Si(OFxn) = 0 Then Exit Function
PushI O, "Some expected ws not found"
For J = 0 To UB(OFxn)
    Fxn = OFxn(J)
    Fx = VzColEqSel(IpxMis, "Fxn", Fxn, "Fx")
    IpxMisi = DwEqSel(IpxMis, "Fxn", Fxn, "L Ws")
    ActWsnn = Termss(FstCol(DwEqE(ActWs, "Fxn", Fxn)))
    '-
    X "Fxn    : " & Fxn
    X "Fx pth : " & Pth(Fx)
    X "Fx fn  : " & Fn(Fx)
    X "Act ws : " & ActWsnn
    X LyzNmDrs("Mis ws : ", IpxMisi)
    PushIAy O, TabAy(XX)
Next
XExWsMis = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XExWsMis) ExWsMis IpxMis ActWs",ExWsMis, ExWsMis, FmtDrs(IpxMis), FmtDrs(ActWs): Stop
End Function

Sub A()
QIde_B_AlignMth.AlignMth
End Sub

Sub ACmdApply()
QIde_B_AlignMth.AlignMthzNm "QXls_Cmd_ApplyFilter", "CmdApply"
End Sub

Sub AU()
QIde_B_AlignMth.AlignMth Upd:=EiUpdAndRpt
End Sub

Sub AUO()
QIde_B_AlignMth.AlignMth Upd:=EiUpdOnly
End Sub

Sub E(Optional Upd As EmUpd)
Dim Mdn$:              Mdn = "QIde_Ens_CModSub"
Dim Mthn$:            Mthn = "EnsCModSubzM"
Dim M As CodeModule: Set M = Md(Mdn)
Dim L&:                  L = MthLnozMM(M, Mthn)
QIde_B_AlignMth.AlignMthzLno M, L, Upd:=Upd
End Sub

Sub FF(Optional Upd As EmUpd)
Dim Mdn$: Mdn = "QIde_Ens_CModSub"
Dim Mthn$: Mthn = "EnsCModSubzM"
Dim M As CodeModule: Set M = Md(Mdn)
Dim L&:                  L = MthLnozMM(M, Mthn)
QIde_B_AlignMth.AlignMthzLno M, L, Upd:=Upd
End Sub

Private Sub Z_ErzLnk()
Brw ErzLnk(Y_LnkImpSrc)
End Sub

Private Function XEiFfnDup(Ipf As Drs) As String()
'Fm  Ipf : L Inpn Ffn IsFx HasFfn @@
Dim Ffn$(): Ffn = StrCol(Ipf, "Ffn")
Dim Dup$(): Dup = AwDup(Ffn)
If Si(Dup) = 0 Then Exit Function
Dim DupD As Drs: DupD = DwIn(Ipf, "Ffn", Dup)
XBox "Ffn Duplicated"
XDrs DupD
XLin
XEiFfnDup = XX
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEiFfnDup) EiFfnDup Ipf",EiFfnDup, EiFfnDup, FmtDrs(Ipf): Stop
End Function

Private Function XEi_AddCol_Pth_Fn(A As Drs) As Drs
Dim ODy()
Dim Dr: For Each Dr In Itr(A.Dy)
    Dim Ffn$: Ffn = Dr(2)
    PushIAy Dr, Array(Pth(Ffn), Fn(Ffn))
    PushI ODy, Dr
Next
XEi_AddCol_Pth_Fn = AddColzFFDy(A, "Pth Fn", ODy)
End Function
Private Function XEiFfnMis(Ipf As Drs) As String()
'Fm  Ipf : L Inpn Ffn IsFx HasFfn @@
If NoReczDrs(Ipf) Then Exit Function
Dim A As Drs: A = DwEq(Ipf, "HasFfn", True) '! L Inp Ffn IsFx HasFfn
Dim B As Drs: B = XEi_AddCol_Pth_Fn(A)
Dim C As Drs: C = SelDrs(B, "L Inpn Pth Fn")
      XEiFfnMis = LyzNmDrs("File missing: ", C, MaxColWdt:=200)

'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEiFfnMis) EiFfnMis Ipf",EiFfnMis, EiFfnMis, FmtDrs(Ipf): Stop
End Function

Private Function XEiInpnDup(Ipf As Drs) As String()
'Fm  Ipf : L Inpn Ffn IsFx HasFfn @@
Dim Inpn$(): Inpn = StrCol(Ipf, "Inpn")
Dim Dup$(): Dup = AwDup(Inpn)
If Si(Dup) = 0 Then Exit Function
Dim DupD As Drs: DupD = DwIn(Ipf, "Inpn", Dup)
XBox "Inpn Duplicated"
XDrs DupD
XLin
XEiInpnDup = XX
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEiInpnDup) EiInpnDup Ipf",EiInpnDup, EiInpnDup, FmtDrs(Ipf): Stop
End Function
Private Function XExTblDup(Ipx As Drs) As String()
'Fm  Ipx : T Fxn Ws Stru ! Inp-Fx which is FxTbl @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XExTblDup) ExTblDup Ipx",ExTblDup, ExTblDup, FmtDrs(Ipx): Stop
End Function
Private Function XExFxnDup(Ipx As Drs) As String()
'Fm  Ipx : T Fxn Ws Stru ! Inp-Fx which is FxTbl @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XExFxnDup) ExFxnDup Ipx",ExFxnDup, ExFxnDup, FmtDrs(Ipx): Stop
End Function
Private Function XExFxnMis(Ipx As Drs, IpfExi As Drs) As String()
'Fm  Ipx    : T Fxn Ws Stru   ! Inp-Fx which is FxTbl
'Fm  IpfExi : L Inpn Ffn IsFx @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XExFxnMis) ExFxnMis Ipx IpfExi",ExFxnMis, ExFxnMis, FmtDrs(Ipx), FmtDrs(IpfExi): Stop
End Function
Private Function XExStruMis(Ipx As Drs, IpsStru$()) As String()
'Fm  Ipx : T Fxn Ws Stru ! Inp-Fx which is FxTbl @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XExStruMis) ExStruMis Ipx IpsStru",ExStruMis, ExStruMis, FmtDrs(Ipx), IpsStru: Stop
End Function
Private Function XEbFbnDup(Ipb As Drs) As String()
'Fm  Ipb : L Fbn Fbtt @@
Dim Ix As Dictionary, IxL%, IxFbn%, Dup, LAy&(), Fbn, Lss, OLss$(), OFbn$(), J%
Set Ix = DiKqIx(Ipb.Fny)
IxL = Ix("L")
IxFbn = Ix("Fbn")
For Each Dup In Itr(AwDup(StrCol(Ipb, "Fbn")))
    LAy = LngAyzColEqSel(Ipb, "Fbn", Dup, "L")
    Fbn = Dup
    Lss = JnSpc(LAy)
    PushI OLss, Lss
    PushI OFbn, Fbn
Next
'===
Dim O$()
For J = 0 To UB(OLss)
    PushI O, FmtQQ("L#(?) Fbn(?) Fbn Duplicated.", OLss(J), OFbn(J))
Next
XEbFbnDup = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEbFbnDup) EbFbnDup Ipb",EbFbnDup, EbFbnDup, FmtDrs(Ipb): Stop
End Function

Private Function XEbFbnMis(Ipb As Drs, IpbFb$()) As String()
'Fm  Ipb : L Fbn Fbtt @@
Dim Ix As Dictionary, IxL%, IxFbn%, Dr, Fbn$, OL&(), L&, OFbn$(), J%, Inpn, O$()
Set Ix = DiKqIx(Ipb.Fny)
IxL = Ix("L")
IxFbn = Ix("Fbn")
For Each Dr In Itr(Ipb.Dy)
    Fbn = Dr(IxFbn)
    If Not HasEle(IpbFb, Fbn) Then
        L = Dr(IxL)
        PushI OL, L
        PushI OFbn, Fbn
    End If
Next
'===
If Si(OL) = 0 Then Exit Function
For J = 0 To UB(OL)
    PushI O, FmtQQ("L#(?) Fbn(?) Fbn is not defined.", OL(J), OFbn(J))
Next
PushI O, vbTab & FmtQQ("Total (?)-Fbn are defined are:", Si(IpbFb))
For Each Inpn In Itr(IpbFb)
    PushI O, vbTab & vbTab & Inpn
Next
XEbFbnMis = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEbFbnMis) EbFbnMis Ipb IpbFb",EbFbnMis, EbFbnMis, FmtDrs(Ipb), IpbFb: Stop
End Function

Private Function XEbTblDup(Ipb As Drs) As String()
'Fm  Ipb : L Fbn Fbtt @@
Dim J&, OL&(), OFbtt$(), Fbtt$, L&, IxL%, IxFbtt%, Ix As Dictionary, B$, Dr, O$()
Set Ix = DiKqIx(Ipb.Fny)
IxL = Ix("L")
IxFbtt = Ix("Fbtt")
For Each Dr In Itr(Ipb.Dy)
    Fbtt = Dr(IxFbtt)
    B = Termss(AwDup(TermAy(Fbtt)))
    If B <> "" Then
        L = Dr(IxL)
        PushI OL, L
        PushI OFbtt, B
    End If
Next
'===
For J = 0 To UB(OL)
    PushI O, FmtQQ("L#(?) Fbtt(?) Fbt Duplicated.", OL(J), OFbtt(J))
Next
XEbTblDup = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEbTblDup) EbTblDup Ipb",EbTblDup, EbTblDup, FmtDrs(Ipb): Stop
End Function

Private Function XEbStruMis(IpbTny$(), IpsStru$()) As String()
'Ret EbStruMis : Use IpbTny stru to find in IpsStru @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEbStruMis) EbStruMis IpbTny IpsStru",EbStruMis, EbStruMis, IpbTny, IpsStru: Stop
End Function
Private Function XEbTblMis(Ipb As Drs) As String()
'Fm  Ipb : L Fbn Fbtt @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEbTblMis) EbTblMis Ipb",EbTblMis, EbTblMis, FmtDrs(Ipb): Stop
End Function

Private Function XEsFldDup(Ips As Drs) As String()
'Fm  Ips : L Stru F Ty E @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEsFldDup) EsFldDup Ips",EsFldDup, EsFldDup, FmtDrs(Ips): Stop
End Function

Private Function XEoNoFxAndNoFb$(Ipx As Drs, Ipb As Drs)
'Fm  Ipb : L Fbn Fbtt @@
If Si(Ipx.Dy) > 0 Then Exit Function
If Si(Ipb.Dy) > 0 Then Exit Function
XEoNoFxAndNoFb = "Both [FxTbl] and [FbTbl] sections are missing (EoNoFxAndNoFb)"
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEoNoFxAndNoFb) EoNoFxAndNoFb Ipx Ipb",EoNoFxAndNoFb, EoNoFxAndNoFb, FmtDrs(Ipx), FmtDrs(Ipb): Stop
End Function

Private Function XEsTyEr() As String()
'Fm:DTyEr@DE?::Drs{ErTy L Stru F E}

Dim O$()
'PushI O, FldEr2 & "Valid Ty are: ...."
'PushIAy O, AddPfxzAy(FmtDrs(DTyEr), vbTab)
XEsTyEr = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEsTyEr) EsTyEr ",EsTyEr, EsTyEr: Stop
End Function

Private Function XIsT1Er(T1$) As Boolean
Select Case T1
Case "NRec", "Inp", "FxTbl", "FbTbl", "Tbl.Where": Exit Function
End Select
If HasPfx(T1, "Stru.") Then
    If IsNm(RmvPfx(T1, "Stru.")) Then Exit Function
End If
XIsT1Er = True
End Function

Private Function XEoHdrEr(Ip As DoLTDH) As String()
'Fm  Ip : L T1 Dta IsHdr @@
Dim OL&(), OT1$(), J%, Dr, T1$, O$()
For Each Dr In Itr(Ip.D.Dy)
    T1 = Dr(1)
    If XIsT1Er(T1) Then
        PushI OL, Dr(0)
        PushI OT1, T1
    End If
Next
'====
If Si(OL) = 0 Then Exit Function
For J = 0 To UB(OL)
'    PushI O, OthEr2 & FmtQQ(" L#(?) T1(?) T1 Error", OL(J), OT1(J))
Next
Push O, vbTab & "Valid T1 are: Inp FxTbl FbTbl Tbl.Where Stru.{Nm}"
XEoHdrEr = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEoHdrEr) EoHdrEr Ip",EoHdrEr, EoHdrEr, FmtDrs(Ip.D): Stop
End Function

Private Function XIpxfMis(Ipxf As Drs, ActWsf As Drs) As Drs
'xfMis
'Fm: @Ipxf::Drs{}
'Ret:*IpxfMis::Drs{}
Dim A As Drs, B As Drs, O As Drs
A = LDrszJn(Ipxf, ActWsf, "Fxn Ws E:F", "Ty:ActTy", "HasF")
B = DwEqE(A, "HasF", False)
O = DrpCol(B, "ActTy")
XIpxfMis = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpxfMis) IpxfMis Ipxf ActWsf",FmtDrs(IpxfMis), FmtDrs(IpxfMis), FmtDrs(Ipxf), FmtDrs(ActWsf): Stop
End Function

Private Function XActWsf(IpxExi As Drs) As Drs
'Fm  IpxExi : IpxExi @@
Dim Dr, IDr, O As Drs, OFny$(), ODy(), F$, Ty$, Fx$, Ws$, IFx%, IWs%
'BrwDrs IpxExi.D: Stop
IFx = IxzAy(IpxExi.Fny, "Fx")
IWs = IxzAy(IpxExi.Fny, "Ws")
For Each Dr In Itr(IpxExi.Dy)
    Fx = Dr(IFx)
    Ws = Dr(IWs)
    For Each IDr In Itr(DoFTyzFxw(Fx, Ws).Dy)
        PushI ODy, AddAy(Dr, IDr)
    Next
Next
OFny = Sy(IpxExi.Fny, "F", "Ty")
O = Drs(OFny, ODy)
XActWsf = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XActWsf) ActWsf IpxExi",FmtDrs(ActWsf), FmtDrs(ActWsf), FmtDrs(IpxExi): Stop
End Function
Private Function XIpbTny(Ipb As Drs) As String()
'Fm  Ipb : L Fbn Fbtt @@
Dim Dr, Fbtt$
For Each Dr In Itr(Ipb.Dy)
    Fbtt = Dr(2)
    PushNDupAy XIpbTny, SyzSS(Fbtt)
Next
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbTny) IpbTny Ipb",IpbTny, IpbTny, FmtDrs(Ipb): Stop
End Function

Private Function XIpxTny(Ipx As Drs) As String()
'Fm  Ipx : T Fxn Ws Stru ! Inp-Fx which is FxTbl @@
Dim Dr
For Each Dr In Itr(Ipx.Dy)
    PushNDup XIpxTny, Dr(1)
Next
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpxTny) IpxTny Ipx",IpxTny, IpxTny, FmtDrs(Ipx): Stop
End Function

Private Function XIpb(Ip As DoLTDH) As Drs
Dim Dr, L&, Dta$, Fbn$, Fbtt$, Dy()
For Each Dr In Itr(DoLDta(Ip, "FbTbl").D.Dy)
    L = Dr(0)
    Dta = Dr(1)
    Fbn = T1(Dta)
    Fbtt = RmvT1(Dta)
    PushI Dy, Array(L, Fbn, Fbtt)
Next
XIpb = DrszFF("L Fbn Fbtt", Dy)
End Function

Private Function XIpf(Ip As DoLTDH) As Drs
'Fm  Ip  : L T1 Dta IsHdr
'Ret     : L Inpn Ffn IsFx HasFfn @@
Dim Dr, Dy(), LTT As Drs, Ix As Dictionary, L&, Inpn$, Ffn$
LTT = DoLTT(Ip, "Inp", "Inpn Ffn").D
For Each Dr In Itr(LTT.Dy)
    L = Dr(0)
    Inpn = Dr(1)
    Ffn = Dr(2)
    PushI Dy, Array(L, Inpn, Ffn, IsFx(Ffn), HasFfn(Ffn))
Next
XIpf = DrszFF("L Inpn Ffn IsFx HasFfn", Dy)
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpf) Ipf Ip",FmtDrs(Ipf), FmtDrs(Ipf), FmtDrs(Ip.D): Stop
End Function

Private Function XIpw(Ip As DoLTDH) As Drs
'Fm  Ip : L T1 Dta IsHdr @@
Dim Dr, L&, Dta$, T$, Bexp$, Dy()
For Each Dr In Itr(DoLDta(Ip, "Tbl.Where").D.Dy)
    L = Dr(0)
    Dta = Dr(1)
    T = T1(Dta)
    Bexp = RmvT1(Dta)
    PushI Dy, Array(L, T, Bexp)
Next
XIpw = DrszFF("L T Bexp", Dy)
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpw) Ipw Ip",FmtDrs(Ipw), FmtDrs(Ipw), FmtDrs(Ip.D): Stop
End Function

Private Function XIpxHasFx(Ipx As Drs, Ipf As Drs) As Drs
'Fm  Ipx      : T Fxn Ws Stru                      ! Inp-Fx which is FxTbl
'Fm  Ipf      : L Inpn Ffn IsFx HasFfn
'Ret IpxHasFx : T Fxn Ws Stru Fx IsFx HasFx HasInp ! Add @@
XIpxHasFx = LDrszJn(Ipx, Ipf, "Fxn:Inpn", "Ffn:Fx IsFx HasFfn:HasFx", "HasInp")
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpxHasFx) IpxHasFx Ipx Ipf",FmtDrs(IpxHasFx), FmtDrs(IpxHasFx), FmtDrs(Ipx), FmtDrs(Ipf): Stop
End Function

Private Function XIpx1(A As DoLTDH) As Drs
'Ret::*Ipx::Drs{L T Fxn Ws Stru}
'BrwDrs A.D: Stop
Dim Dr, L&, Dta$, Dy(), C As Drs, IxT%, IxL%, IxFxnDotWs%, IxStru%, _
Stru$, B$, T$, Ix As Dictionary, Fxn$, Ws$
C = DoLTTT(A, "FxTbl", "T FxnDotWs Stru").D
Set Ix = DiKqIx(C.Fny)
IxT = Ix("T")
IxL = Ix("L")
IxFxnDotWs = Ix("FxnDotWs")
IxStru = Ix("Stru")
For Each Dr In Itr(C.Dy)
    L = Dr(IxL)
    T = Dr(IxT)
    B = Dr(IxFxnDotWs)
    Stru = Dr(IxStru)
    Fxn = BefDotOrAll(B)
    Ws = AftDot(B)
    If Fxn = "" Then Fxn = T
    If Ws = "" Then Ws = "Sheet1"
    If Stru = "" Then Stru = T
    PushI Dy, Array(L, T, Fxn, Ws, Stru)
Next
XIpx1 = DrszFF("L T Fxn Ws Stru", Dy)
'BrwDrs XIpx.D: Stop
End Function

'================================================
Private Function XEsSDup(IpsHdStru$()) As String()
'Fm  IpsHdStru :  ! the stru coming from the Ips hd @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEsSDup) EsSDup IpsHdStru",EsSDup, EsSDup, IpsHdStru: Stop
End Function
Private Function XEsSMis(IpsHdStru$(), IpbxStru$()) As String()
'Fm  IpsHdStru :  ! the stru coming from the Ips hd
'Fm  IpbxStru  :  ! the stru comming from Ipb and Ipx @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEsSMis) EsSMis IpsHdStru IpbxStru",EsSMis, EsSMis, IpsHdStru, IpbxStru: Stop
End Function
Private Function XEsSExa(IpsHdStru$(), IpbxStru$()) As String()
'Fm  IpsHdStru :  ! the stru coming from the Ips hd
'Fm  IpbxStru  :  ! the stru comming from Ipb and Ipx @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEsSExa) EsSExa IpsHdStru IpbxStru",EsSExa, EsSExa, IpsHdStru, IpbxStru: Stop
End Function
Private Function XEsSNoFld(IpsHdStru$(), IpsStru$()) As String()
'Fm  IpsHdStru :  ! the stru coming from the Ips hd @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEsSNoFld) EsSNoFld IpsHdStru IpsStru",EsSNoFld, EsSNoFld, IpsHdStru, IpsStru: Stop
End Function
Private Function XLsszWhT$(Wh As Drs, T)
'Fm:Wh@Ipw::Drs{L T Bexp}
Dim O&(), Dr
For Each Dr In Itr(Wh.Dy)
    If Dr(1) = T Then
        Push O, Dr(0)
    End If
Next
XLsszWhT = JnSpc(O)
End Function
Private Function XEwTblExa(Ipw As Drs, Tny$()) As String()
'Ret EwTblExa : tbl.wh is more @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEwTblExa) EwTblExa Ipw Tny",EwTblExa, EwTblExa, FmtDrs(Ipw), Tny: Stop
End Function
Private Function XEwTblDup(Ipw As Drs) As String()
'Fm:Wh@Ipw::Drs{L T Bexp}
Dim OLss$(), OT$(), J%, T, Dr, DupTny$(), Dup, O$()
DupTny = AwDup(StrCol(Ipw, "T"))
For Each Dup In Itr(DupTny)
    PushI OLss, XLsszWhT(Ipw, Dup)
    PushI OT, Dup
Next
'===
If Si(OLss) = 0 Then Exit Function
For J = 0 To UB(OLss)
    PushI O, FmtQQ("L#(?) Tbl(?) Tbl are dup", OLss(J), OT(J))
Next
XEwTblDup = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEwTblDup) EwTblDup Ipw",EwTblDup, EwTblDup, FmtDrs(Ipw): Stop
End Function
Private Function XEwTblMis(Ipw As Drs, Tny$()) As String()
'Fm:Wh@Ipw::Drs{L T Bexp}
Dim OL&(), OT$(), J%, T, Dr, O$()
For Each Dr In Itr(Ipw.Dy)
    T = Dr(1)
    If Not HasEle(Tny, T) Then
        PushI OL, Dr(0)
        PushI OT, T
    End If
Next
'====
If Si(OL) = 0 Then Exit Function
For J = 0 To UB(OL)
    PushI O, FmtQQ("L#(?) Tbl(?) is not defined.", OL(J), OT(J))
Next
PushI O, vbTab & "Defined tables are:"

For Each T In Itr(Tny)
    PushI O, vbTab & vbTab & T
Next
XEwTblMis = O
End Function
Private Function XEwBexpEmp(Ipw As Drs) As String()
'Ret            : with tbl nm but no Bexp @@
Dim J%, OL&(), OT$(), O$()
'Fm : Wh@Ipw::Drs{L T Bexp}
Dim Dr, L&, T$, Bexp$
For Each Dr In Itr(Ipw.Dy)
    Bexp = Dr(2)
    If Bexp = "" Then
        L = Dr(0)
        T = Dr(1)
        PushI OL, L
        PushI OT, T
    End If
Next
'===
For J = 0 To UB(OL)
    PushI O, FmtQQ("L#(?) Tbl(?) has no Bexp", OL(J), OT(J))
Next
XEwBexpEmp = O
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XEwBexpEmp) EwBexpEmp Ipw",EwBexpEmp, EwBexpEmp, FmtDrs(Ipw): Stop
End Function

Private Function DoAddInpIfEr(E$(), InpFilSrc$(), LnkImpSrc$()) As String()
If Si(E) = 0 Then Exit Function
Dim O$(): O = E
PushIAy O, LyzNmLy("InpFilSrc", InpFilSrc, EiBeg1)
PushIAy O, LyzNmLy("LnkImpSrc", LnkImpSrc, EiBeg1)
DoAddInpIfEr = O
End Function
Private Function XIpsFx(Ipx As Drs) As String()
'Fm  Ipx   : T Fxn Ws Stru   ! Inp-Fx which is FxTbl
'Ret IpsFx : IpsFx::Sy{Stru} ! the stru used by Fx @@
XIpsFx = AwDist(StrCol(Ipx, "Stru"))
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpsFx) IpsFx Ipx",IpsFx, IpsFx, FmtDrs(Ipx): Stop
End Function
Private Function XIps1(Ips As Drs, IpsFx$()) As Drs
'Fm  Ips   : L Stru F Ty E
'Fm  IpsFx : IpsFx::Sy{Stru}       ! the stru used by Fx
'Ret Ips1  : L Stru F Ty E IsIpsFx @@
Dim Dr, Stru$, IxStru%, ODy()
IxStru = IxzAy(Ips.Fny, "Stru")
For Each Dr In Itr(Ips.Dy)
    Stru = Dr(IxStru)
    PushI Dr, HasEle(IpsFx, Stru)
    PushI ODy, Dr
Next
XIps1 = Drs(AddSS(Ips.Fny, "IsIpsFx"), ODy)
'BrwDrs XIps1: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIps1) Ips1 Ips IpsFx",FmtDrs(Ips1), FmtDrs(Ips1), FmtDrs(Ips), IpsFx: Stop
End Function

Private Function XIps(Ip As DoLTDH) As Drs
'Fm  Ip  : L T1 Dta IsHdr
'Ret     : L Stru F Ty E @@
Dim B As Drs, Dr, L&, T1$, Dta$, F$, Ty$, E$, Stru$, ODy()
B = DoLDtazT1Pfx(Ip, "Stru.").D
For Each Dr In Itr(B.Dy)
    L = Dr(0)
    T1 = Dr(1)
    Dta = Dr(2)
    Stru = T1
    F = ShfT1(Dta)
    Ty = ShfT1(Dta)
    E = RmvSqBkt(Dta)
    PushI ODy, Array(L, Stru, F, Ty, E)
Next
XIps = DrszFF("L Stru F Ty E", ODy)
'BrwDrs XIps: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIps) Ips Ip",FmtDrs(Ips), FmtDrs(Ips), FmtDrs(Ip.D): Stop
End Function
Private Property Get Y_InpFilSrc() As String()
X "Inp"
X " DutyPay C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
X " ZHT0    C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\Pricing report(ForUpload).xls"
X " MB52    C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
X " Uom     C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\sales text.xlsx"
X " GLBal   C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\DutyPrepayGLTot.xlsx"
Y_InpFilSrc = XX
End Property

Private Property Get Y_LnkImpSrc() As String()
X Y_InpFilSrc
X "FbTbl"
X "--  Fbn T.."
X " DutyPay Permit PermitD"
X "FxTbl T  FxNm.Wsn  Stru"
X " ZHT086  ZHT0.8600 ZHT0"
X " ZHT087  ZHT0.87011 ZHT0"
X " MB52"
X " Uom"
X " GLBal"
X "Tbl.Where"
X " MB52 Plant='8601' and [Storage Location] in ('0002','')"
X " Uom  Plant='8601'"
X "Stru.Permit"
X " Permit"
X " PermitNo"
X " PermitDate"
X " PostDate"
X " NSku"
X " Qty"
X " Tot"
X " GLAc"
X " GLAcName"
X " BankCode"
X " ByUsr"
X " DteCrt"
X " DteUpd"
X "Stru.PermitD"
X " PermitD"
X " Permit"
X " Sku"
X " SeqNo"
X " Qty"
X " BchNo"
X " Rate"
X " Amt"
X " DteCrt"
X " DteUpd"
X "Stru.ZHT0"
X " Sku       Txt Material    "
X " CurRateAc Dbl [     Amount]"
X " VdtFm     Txt Valid From  "
X " VdtTo     Txt Valid to    "
X " HKD       Txt Unit        "
X " Per       Txt per         "
X " CA_Uom    Txt Uom         "
X "Stru.MB52"
X " Sku1    Txt Material          "
X " Sku    Txt Material          "
X " Whs    Txt Plant             "
X " Loc    Txt Storage Location  "
X " BchNo  Txt Batch             "
X " QInsp  Dbl In Quality Insp#  "
X " QUnRes Dbl UnRestricted      "
X " QBlk   Dbl Blocked           "
X " VInsp  Dbl Value in QualInsp#"
X " VUnRes Dbl Value Unrestricted"
X " VBlk   Dbl Value BlockedStock"
X " VBlk2  Dbl Value BlockedStock1"
X " VBlk1  Dbl Value BlockedStock2"
X "Stru.Uom"
X " Sc_U    Txt SC "
X " Topaz   Txt Topaz Code "
X " ProdH   Txt Product hierarchy"
X " Sku     Txt Material            "
X " Des     Txt Material Description"
X " AC_U    Txt Unit per case       "
X " SkuUom  Txt Base Unit of Measure"
X " BusArea Txt Business Area       "
X "Stru.GLBal"
X " BusArea Txt Business Area Code"
X " GLBal   Dbl                   "
X "Stru.SkuTaxBy3rdParty"
X " SkuTaxBy3rdParty "
X "Stru.SkuNoLongerTax"
X " SkuNoLongerTax"
X "NRec"
X " GT 0 *All"
Y_LnkImpSrc = XX
End Property


Private Function XIpxStru(Ipx As Drs) As String()
'Fm  Ipx     : T Fxn Ws Stru ! Inp-Fx which is FxTbl
'Ret IpxStru :               ! the stru coming Ipx @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpxStru) IpxStru Ipx",IpxStru, IpxStru, FmtDrs(Ipx): Stop
End Function

Private Function XIpsHdStru(Ips As Drs) As String()
'Fm  Ips       : L Stru F Ty E
'Ret IpsHdStru :               ! the stru coming from the Ips hd @@
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpsHdStru) IpsHdStru Ips",IpsHdStru, IpsHdStru, FmtDrs(Ips): Stop
End Function

Private Function XIpbxStru$()
End Function


'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
Private Function XIpbStru() As String()
'Insp "QDao_Lnk_ErzLnk.ErzLnk", "Inspect", "Oup(XIpbStru) IpbStru ",IpbStru, IpbStru: Stop
End Function
