Attribute VB_Name = "QDao_Lnk_ErzLnk"
Option Compare Text
Option Explicit
Const InpEr1$ = "#InpEr1-InpnDup. "
Const InpEr2$ = "#InpEr2-FfnDup. "
Const InpEr3$ = "#InpEr3-FfnMis. "
Const FxEr1$ = "#FxEr1-TblDup. "
Const FxEr2$ = "#FxEr2-FxnDup. "
Const FxEr3$ = "#FxEr3-FxnMis. "
Const FxEr4$ = "#FxEr4-WsMis. "
Const FxEr5$ = "#FxEr5-WsMisFld. "
Const FxEr6$ = "#FxEr6-WsMisFldTy. "
Const FxEr7$ = "#FxEr7-StruMis. "
Const FbEr1$ = "#FbEr1-FbnDup. "
Const FbEr2$ = "#FbEr2-FbnMis. "
Const FbEr3$ = "#FbEr3-TblDup. "
Const FbEr4$ = "#FbEr4-TblMis. "
Const FbEr5$ = "#FbEr5-StruMis. "
Const StruEr1$ = "#StruEr1-Mis."
Const StruEr2$ = "#StruEr1-Dup."
Const StruEr3$ = "#StruEr1-Exa."
Const StruEr4$ = "#StruEr1-NoFld."
Const FldEr1$ = "#FldEr1-FldDup."
Const FldEr2$ = "#FldEr2-TyEr."
Const OthEr1$ = "#InpEr1-NoFxAndNoFb. "
Const OthEr2$ = "#OthEr2-SectEr. "
Const WhEr1$ = "#WhEr1-TblDup. "
Const WhEr2$ = "#WhEr2-TblMis. "
Const WhEr3$ = "#WhEr3-BexprEmp. "


Private Function XIpfExi(Ipf As Drs) As Drs
XIpfExi = DrswColEqExlEqCol(Ipf, "HasFfn", True)
End Function
Private Function XIpfMis(Ipf As Drs) As Drs
XIpfMis = DrswColEqExlEqCol(Ipf, "HasFfn", False)
End Function
Private Function XIpbFb(Ipf As Drs) As String()
XIpbFb = StrColzDrswColEqSel(Ipf, "IsFx", False, "Inpn")
End Function
Function ErzLnk(LnkImpSrc$()) As String()
'Fm: *InpFilSrc::SSAy{Inpn Ffn}
'Fm: *LnkImpSrc::IndentedLy{
'                  TblFx: {Fxt} [{Fxn}[.{Wsn}]] [{Stru}]
'                  TblFb: {Fbn} {Fbtt}
'                  Stru.{Stru}: {F} [{Ty}] [{Extn}]
'                  Tbl.Where: {T} {Bexpr}
'                }
'--
' #Act
' #Ip  : Inp
'  Ipb : Ip-Fb
'  Ipf : Ip-Fx
'  Ipw : Ip-Wh
'  Ipf : Ip-Fil
'  Ips : Ip-Stru
'  Ipn : Ip-Nrec
' #E  : Er
'  Eb : E-Fb
'  Ex : E-Fx
'  Ew : E-Wh
'  Ef : E-FEl
'  Es : E-Stru
'  En : E-Nrec
'  Eo : E-Oth
' *Eb : E-Fb
'--
Const CmPfx$ = "X"
Dim Ip     As DLTDH:     Ip = DLTDH(LnkImpSrc)                      ' L T1 Dta IsHdr
Dim Ipf    As Drs:      Ipf = XIpf(Ip)                              ' L Inpn Ffn IsFx HasFfn
Dim IpfFx  As Drs:    IpfFx = DrswColEqExlEqCol(Ip.D, "IsFx", True)
Dim IpfExi As Drs:   IpfExi = XIpfExi(Ipf)                          ' L Inpn Ffn IsFx
Dim IpfMis As Drs:   IpfMis = XIpfMis(Ipf)                          ' L Inpn Ffn IsFx
Dim IpbFb$():         IpbFb = XIpbFb(Ipf)
Dim Ipw    As Drs:      Ipw = XIpw(Ip)

Dim Ipx      As Drs:      Ipx = XIpx1(Ip)                                       ' T Fxn Ws Stru                        ! Inp-Fx which is FxTbl
Dim IpxHasFx As Drs: IpxHasFx = XIpxHasFx(Ipx, Ipf)                             ' T Fxn Ws Stru Fx IsFx HasFx HasInp   ! Add
Dim D1       As Drs:       D1 = DrswColEqExlEqCol(IpxHasFx, "IsFx", True)
Dim D2       As Drs:       D2 = DrswColEqExlEqCol(D1, "HasFx", True)
Dim D3       As Drs:       D3 = DrswColEqExlEqCol(D2, "HasInp", True)
Dim IpxExiFx As Drs: IpxExiFx = D3                                              ' L T Fxn Ws Stru Fx
Dim ActWs    As Drs:    ActWs = XActWs(IpxExiFx)                                ' Fxn Ws
Dim IpxHasWs As Drs: IpxHasWs = AddColzExiB(IpxExiFx, ActWs, "Fxn Ws", "HasWs") ' T Fxn Ws Stru Fx HasFx HasWs
Dim IpxExi   As Drs:   IpxExi = DrswColEqExlEqCol(IpxHasWs, "HasWs", True)      ' IpxExi   ::Drs{}
Dim IpxMis   As Drs:   IpxMis = DrswColEqExlEqCol(IpxHasWs, "HasWs", False)

Dim IpsFx$():      IpsFx = XIpsFx(Ipx)             ' IpsFx::Sy{Stru}         ! the stru used by Fx
Dim Ips  As Drs:     Ips = XIps(Ip)                ' L Stru F Ty E
Dim Ips1 As Drs:    Ips1 = XIps1(Ips, IpsFx)       ' L Stru F Ty E IsIpsFx
Dim IpsStru$():  IpsStru = StrColzDrs(Ips, "Stru")

Dim Ipb As Drs:    Ipb = DLTT(Ip, "FbTbl", "Fbn Fbtt").D ' Ipb::Drs{L Fbn Fbtt}
Dim IpxTny$():  IpxTny = XIpxTny(Ipx)
Dim IpbTny$():  IpbTny = XIpbTny(Ipb)
Dim Tny$():        Tny = Sy(IpxTny, IpbTny)

Dim ActWsf  As Drs:  ActWsf = XActWsf(IpxExi)
Dim Ipxf    As Drs:    Ipxf = XIpxf(IpxExi, Ips)
Dim IpxfMis As Drs: IpxfMis = XIpxfMis(Ipxf, ActWsf)
'== Error===============================================================================================================
Dim EiInpnDup$(): EiInpnDup = XEiInpnDup(Ipf)
Dim EiFfnDup$():   EiFfnDup = XEiFfnDup(Ipf)
Dim EiFfnMis$():   EiFfnMis = XEiFfnMis(IpfMis)
Dim I$():                 I = Sy(EiInpnDup, EiFfnDup, EiFfnMis)

Dim ExTblDup$():         ExTblDup = XExTblDup(Ipx)
Dim ExFxnDup$():         ExFxnDup = XExFxnDup(Ipx)
Dim ExFxnMis$():         ExFxnMis = XExFxnMis(Ipx, IpfExi)
Dim ExWsMis$():           ExWsMis = XExWsMis(IpxMis, ActWs)
Dim ExWsMisFld$():     ExWsMisFld = XExWsMisFld(IpxfMis, ActWsf)
Dim ExWsMisFldTy$(): ExWsMisFldTy = XExWsMisFldTy(Ipxf, ActWsf)
Dim ExStruMis$():       ExStruMis = XExStruMis(Ipx, IpsStru)
Dim X$():                       X = Sy(ExFxnDup, ExFxnMis, ExWsMis, ExWsMisFld, ExWsMisFldTy)

Dim EbFbnDup$():   EbFbnDup = XEbFbnDup(Ipb)
Dim EbFbnMis$():   EbFbnMis = XEbFbnMis(Ipb, IpbFb)
Dim EbTblDup$():   EbTblDup = XEbTblDup(Ipb)
Dim EbTblMis$():   EbTblMis = XEbTblMis(Ipb)
Dim EbStruMis$(): EbStruMis = XEbStruMis(IpbTny, IpsStru)                           ' Use IpbTny stru to find in IpsStru
Dim B$():                 B = Sy(EbFbnDup, EbFbnMis, EbTblDup, EbTblMis, EbStruMis)

Dim IpxStru$():      IpxStru = IpxStru
Dim IpsHead As Drs:  IpsHead = IpsHead
Dim IpbxStru$():    IpbxStru = IpbxStru

Dim EsSDup$():     EsSDup = XEsSDup(IpsHead)
Dim EsSMis$():     EsSMis = XEsSMis(IpsHead, IpbxStru)
Dim EsSExa$():     EsSExa = XEsSExa(IpsHead, IpbxStru)
Dim EsSNoFld$(): EsSNoFld = XEsSNoFld(IpsHead, IpsStru)
Dim EsFldDup$(): EsFldDup = XEsFldDup(Ips)
Dim EsTyEr$():     EsTyEr = XEsTyEr(XDTyEr(Ips1))
Dim S$():               S = Sy(EsSDup, EsSMis, EsSExa, EsSNoFld, EsFldDup, EsTyEr)

Dim EwTblDup$():     EwTblDup = XEwTblDup(Ipw)
Dim EwTblExa$():     EwTblExa = XEwTblExa(Ipw, Tny)                ' tbl.wh is more
Dim EwBexprEmp$(): EwBexprEmp = XEwBexprEmp(Ipw)                   ' with tbl nm but no bexpr
Dim W$():                   W = Sy(EwTblDup, EwTblExa, EwBexprEmp)

Dim EoNoFxAndNoFb$: EoNoFxAndNoFb = XEoNoFxAndNoFb(Ipx, Ipb)
Dim EoHdrEr$():           EoHdrEr = XEoHdrEr(Ip)
Dim O$():                       O = Sy(EoNoFxAndNoFb, EoHdrEr)

ErzLnk = Sy(I, X, B, S, W, O)
End Function
Private Function XIpx(Ip As DLTDH) As Drs

End Function

Private Function XIpfFx(Ip As DLTDH) As Drs

End Function

Private Function XIpxf(IpxExi As Drs, Ips As Drs) As Drs
Dim O As Drs, Dr, E$, IxF%, IE%, J&
O = JnDrs(IpxExi, Ips, "Stru", "F Ty E")
IxF = IxzAy(O.Fny, "F")
IE = IxzAy(O.Fny, "E")
For Each Dr In Itr(O.Dry)
    E = Dr(IE)
    If E = "" Then
        Dr(IE) = Dr(IxF)
        O.Dry(J) = Dr
    End If
    J = J + 1
Next
XIpxf = O
'BrwDrs3 O, IpxExi, Ips, NN:="Ipxf IpxExi Ips": Stop
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
Dim Ty$, L&, Stru$, F$, E$, IsIpsFx As Boolean, ODry()
Fny = Ips1.Fny
IL = IxzAy(Fny, "L")
IStru = IxzAy(Fny, "Stru")
IxF = IxzAy(Fny, "F")
ITy = IxzAy(Fny, "Ty")
IE = IxzAy(Fny, "E")
IxIsIpsFx = IxzAy(Fny, "IsIpsFx")
For Each Dr In Itr(Ips1.Dry)
    Ty = Dr(ITy)
    IsIpsFx = Dr(IxIsIpsFx)
    L = Dr(IL)
    Stru = Dr(IStru)
    F = Dr(IxF)
    E = Dr(IE)
    If Not XIsShtTy(Ty, IsIpsFx) Then
        PushI ODry, Array(Ty, L, Stru, F, E, IsIpsFx)
    End If
Next
XDTyEr = DrszFF("TyEr L Stru F E IsIpsFx", ODry)
End Function

Private Function WActTbl(IpfFb As Drs) As Drs
Dim Dr, T, J%, IFbn$, IFb$, Dry()
For Each Dr In Itr(IpfFb.Dry)
    IFbn = Dr(1)
    IFb = Dr(2)
    For Each T In Tni(Db(IFb))
        PushI Dry, Array(IFbn, T)
    Next
Next
WActTbl = DrszFF("Fbn T", Dry)
End Function

Private Function XActWs(IpxExiFx As Drs) As Drs
'Fm :IpxExi::Drs{L T Fxn Ws Stru Fx}
'Ret:*WsAct::Drs{Fxn Ws}
Dim A As Drs, Dr, Fxn$, Fx$, Wsn, Dry()
'A = DrswIpst(IpxExiFx, "Fxn Fx")
For Each Dr In Itr(A.Dry)
    Fxn = Dr(0)
    Fx = Dr(1)
    For Each Wsn In Wny(Fx)
        PushI Dry, Array(Fxn, Wsn)
    Next
Next
XActWs = DrszFF("Fxn Ws", Dry)
'BrwDrs XActWs: Stop
End Function

Function AddFF(Fny$(), FF$) As String()
AddFF = Sy(Fny, SyzSS(FF))
End Function

Private Function XExWsMisFld(IpxfMis As Drs, ActWsf As Drs) As String()
'Fm : @IpxfMis :: Drs{L Ws Fxn Fx}
'Fm : @ActWsf:: SSAy{Fxn Ws F Ty}
If NoReczDrs(IpxfMis) Then Exit Function
Dim OFx$(), OFxn$(), OWs$(), O$(), Fxn, Fx$, Ws$, Mis As Drs, Act As Drs, J%, O1$()
AsgCol IpxfMis, "Fxn Fx Ws", OFxn, OFx, OWs
'====
PushI O, "Some columns in ws is missing"
For Each Fxn In OFxn
    Fxn = OFxn(J)
    Fx = OFx(J)
    Ws = OWs(J)
    Mis = DrswCCCEqExlEqCol(IpxfMis, "Fxn Fx Ws", Fxn, Fx, Ws)
    Act = DrswCCCEqExlEqCol(ActWsf, "Fxn Fx Ws", Fxn, Fx, Ws)
    '-
    
    Erase XX
    X "Fxn    : " & Fxn
    X "Fx pth : " & Pth(Fx)
    X "Fx fn  : " & Fn(Fx)
    X "Ws     : " & Ws
    X LyzNmDrs("Mis col: ", Mis)
    X LyzNmDrs("Act col: ", Act)
    PushIAy O, TabAy(XX)
    J = J + 1
Next
Erase XX
XExWsMisFld = O
End Function
Private Function XExWsMisFldTy(Ipxf As Drs, ActWsf As Drs) As String()
'Fm : @IpxFld:: Drs{Fxn Ws Stru Ipxf Ty Fx} ! Where HasFx and HasWs and Not HasFld
'Fm : @WsActFld::Drs{Fxn Ws Ipxf Ty}
'Dim OFxn$(), J%, Fxn$, Fx$, Act$(), Lno&(), Ws$(), ActWsf()
'OFxn = AywDist(StrColzDrs(Ipx4, "Fxn"))
''====
'If Si(OFxn) = 0 Then Exit Function
'PushI XExWsMis, "Some expected ws not found"
'For J = 0 To UB(OFxn)
'    Fxn = OFxn(J)
'    Fx = ValzDrswColEqSel(Ipx4, "Fxn", Fxn, "Fx")
'    ActWsf = DrswColEqSel(Ipx4, "Fxn", Fxn, "L Ws").Dry
'    Lno = LngAyzDryC(ActWsf, 0)
'    Ws = SyzDryC(ActWsf, 1)
'
'    Act = RmvT1zAy(AywT1(WsAct, Fxn)) '*WsActPerFxn::Sy{WsAct}
'    PushIAy XExWsMis, XMisWs_OneFx(Fxn, Fx, Lno, Ws, Act)
'Next
End Function

Private Function XExWsMis(IpxMis As Drs, ActWs As Drs) As String()
Dim OFxn$(), J%, Fxn$, Fx$, Act$(), Lno&(), Ws$(), ActWsnn$, IpxMisi As Drs, O$()
OFxn = AywDist(StrColzDrs(IpxMis, "Fxn"))
'====
If Si(OFxn) = 0 Then Exit Function
PushI O, "Some expected ws not found"
For J = 0 To UB(OFxn)
    Fxn = OFxn(J)
    Fx = ValzDrswColEqSel(IpxMis, "Fxn", Fxn, "Fx")
    IpxMisi = DrswColEqSel(IpxMis, "Fxn", Fxn, "L Ws")
    ActWsnn = TermLin(FstCol(DrswColEqExlEqCol(ActWs, "Fxn", Fxn)))
    '-
    Erase XX
    X "Fxn    : " & Fxn
    X "Fx pth : " & Pth(Fx)
    X "Fx fn  : " & Fn(Fx)
    X "Act ws : " & ActWsnn
    X LyzNmDrs("Mis ws : ", IpxMisi)
    PushIAy O, TabAy(XX)
Next
Erase XX
XExWsMis = O
End Function

Private Sub ZZ_ErzLnk()
Brw ErzLnk(Y_LnkImpSrc)
End Sub

Private Function XEiFfnDup(Ipf As Drs) As String()
Dim OLss$(), OFfn$(), Dr, L&, Ffn$, Dup, LAy&(), J%, O$()
For Each Dup In Itr(AywDup(StrColzDrs(Ipf, "Ffn")))
    LAy = LngAyzDrswColEqSel(Ipf, "Ffn", Dup, "L")
    PushI OLss, JnSpc(LAy)
    PushI OFfn, Dup
Next
'=====
If Si(OLss) = 0 Then Exit Function
For J = 0 To UB(OLss)
    PushI O, InpEr2 & FmtQQ("L#(?) Ffn(?) Ffn is duplicated", OLss(J), OFfn(J))
Next
XEiFfnDup = O
End Function

Private Function XEiFfnMis(IpfMis As Drs) As String()
If NoReczDry(IpfMis.Dry) Then Exit Function
Dim O$()
O = LyzNmDrs(InpEr1 & " file missing: ", IpfMis, MaxColWdt:=200)
XEiFfnMis = O
End Function

Private Function XEiInpnDup(Ipf As Drs) As String()
'Fm: @Ipf::Drs{L Inpn Inpn IsFx}
Dim OLss$(), OInpn$(), Dr, L&, Inpn$, Dup, LAy&(), J%, O$()
For Each Dup In Itr(AywDup(StrColzDrs(Ipf, "Inpn")))
    LAy = LngAyzDrswColEqSel(Ipf, "Inpn", Dup, "L")
    PushI OLss, JnSpc(LAy)
    PushI OInpn, Dup
Next
'=====
If Si(OLss) = 0 Then Exit Function
For J = 0 To UB(OLss)
    PushI O, InpEr1 & FmtQQ("L#(?) Inpn(?) Inpn is duplicated", OLss(J), OInpn(J))
Next
XEiInpnDup = O
End Function
Private Function XExTblDup(Ipx As Drs) As String()

End Function
Private Function XExFxnDup(Ipx As Drs) As String()

End Function
Private Function XExFxnMis(Ipx As Drs, IpfExi As Drs) As String()

End Function
Private Function XExStruMis(Ipx As Drs, IpsStru$()) As String()

End Function
Private Function XEbFbnDup(Ipb As Drs) As String()
'Fm: Fb@Ipb::Drs{L Fbn Fbtt}
Dim Ix As Dictionary, IxL%, IxFbn%, Dup, LAy&(), Fbn, Lss, OLss$(), OFbn$(), J%
Set Ix = DiczAyIx(Ipb.Fny)
IxL = Ix("L")
IxFbn = Ix("Fbn")
For Each Dup In Itr(AywDup(StrColzDrs(Ipb, "Fbn")))
    LAy = LngAyzDrswColEqSel(Ipb, "Fbn", Dup, "L")
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
End Function

Private Function XEbFbnMis(Ipb As Drs, IpbFb$()) As String()
'Fm: @FbInpAy::Sy{Inpn}
'Fm: Fb@Ipb::Drs{L Fbn Fbtt}
Dim Ix As Dictionary, IxL%, IxFbn%, Dr, Fbn$, OL&(), L&, OFbn$(), J%, Inpn, O$()
Set Ix = DiczAyIx(Ipb.Fny)
IxL = Ix("L")
IxFbn = Ix("Fbn")
For Each Dr In Itr(Ipb.Dry)
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
End Function

Private Function XEbTblDup(Ipb As Drs) As String()
'Fm: Ipb@Ipb::Drs{L Fbn Fbtt}
Dim J&, OL&(), OFbtt$(), Fbtt$, L&, IxL%, IxFbtt%, Ix As Dictionary, B$, Dr, O$()
Set Ix = DiczAyIx(Ipb.Fny)
IxL = Ix("L")
IxFbtt = Ix("Fbtt")
For Each Dr In Itr(Ipb.Dry)
    Fbtt = Dr(IxFbtt)
    B = TermLin(AywDup(TermAy(Fbtt)))
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
End Function

Private Function XEbStruMis(IpbTny$(), IpsStru$()) As String()

End Function
Private Function XEbTblMis(Ipb As Drs) As String()

End Function

Private Function XEsFldDup(Ips As Drs) As String()

End Function

Private Function XEoNoFxAndNoFb$(Ipx As Drs, Ipb As Drs)
If Si(Ipx.Dry) > 0 Then Exit Function
If Si(Ipb.Dry) > 0 Then Exit Function
XEoNoFxAndNoFb = OthEr1 & "Both [FxTbl] and [FbTbl] sections are missing"
End Function

Private Function XEsTyEr(DTyEr As Drs) As String()
'Fm:DTyEr@DE?::Drs{ErTy L Stru F E}
If NoReczDrs(DTyEr) Then Exit Function
Dim O$()
PushI O, FldEr2 & "Valid Ty are: ...."
PushIAy O, AddPfxzAy(FmtDrs(DTyEr), vbTab)
XEsTyEr = O
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

Private Function XEoHdrEr(A As DLTDH) As String()
Dim OL&(), OT1$(), J%, Dr, T1$, O$()
For Each Dr In Itr(A.D.Dry)
    T1 = Dr(1)
    If XIsT1Er(T1) Then
        PushI OL, Dr(0)
        PushI OT1, T1
    End If
Next
'====
If Si(OL) = 0 Then Exit Function
For J = 0 To UB(OL)
    PushI O, OthEr2 & FmtQQ(" L#(?) T1(?) T1 Error", OL(J), OT1(J))
Next
Push O, vbTab & "Valid T1 are: Inp FxTbl FbTbl Tbl.Where Stru.{Nm}"
XEoHdrEr = O
End Function

Private Function XIpxfMis(Ipxf As Drs, ActWsf As Drs) As Drs
'xfMis
'Fm: @Ipxf::Drs{}
'Ret:*IpxfMis::Drs{}
Dim A As Drs, B As Drs, O As Drs
A = LJnDrs(Ipxf, ActWsf, "Fxn Ws E:F", "Ty:ActTy", "HasF")
B = DrswColEqExlEqCol(A, "HasF", False)
O = DrpColzDrs(B, "ActTy")
'BrwDrs4 Ipxf, ActWsf, ActWsf, O: Stop
XIpxfMis = O
End Function

Private Function XActWsf(IpxExi As Drs) As Drs
'Fm : IpxExi@IpxExi::Drs{L T Fxn Ws Stru Fx}
'Ret::*ActWsf::Drs{?? Fxn Ws F Ty}
Dim Dr, IDr, O As Drs, OFny$(), ODry(), F$, Ty$, Fx$, Ws$, IFx%, IWs%
'BrwDrs IpxExi.D: Stop
IFx = IxzAy(IpxExi.Fny, "Fx")
IWs = IxzAy(IpxExi.Fny, "Ws")
For Each Dr In Itr(IpxExi.Dry)
    Fx = Dr(IFx)
    Ws = Dr(IWs)
    For Each IDr In Itr(DFTyzFxw(Fx, Ws).Dry)
        PushI ODry, AddAy(Dr, IDr)
    Next
Next
OFny = Sy(IpxExi.Fny, "F", "Ty")
O = Drs(OFny, ODry)
XActWsf = O
'BrwDrs2 IpxExi, O, NN:="IpxExi ActWsf": Stop
End Function
Function XIpbTny(Ipb As Drs) As String()
'Fm: Ipb::Drs{L Fbn Fbtt}
Dim Dr, Fbtt$
For Each Dr In Itr(Ipb.Dry)
    Fbtt = Dr(2)
    PushNoDupAy XIpbTny, SyzSS(Fbtt)
Next
End Function

Function XIpxTny(Ipx As Drs) As String()
'Fm: @Ipx::Drs{L T Fxn Ws Stru}
Dim Dr
For Each Dr In Itr(Ipx.Dry)
    PushNoDup XIpxTny, Dr(1)
Next
End Function

Private Function XIpb(Ip As DLTDH) As Drs
Dim Dr, L&, Dta$, Fbn$, Fbtt$, Dry()
For Each Dr In Itr(DLDta(Ip, "FbTbl").D.Dry)
    L = Dr(0)
    Dta = Dr(1)
    Fbn = T1(Dta)
    Fbtt = RmvT1(Dta)
    PushI Dry, Array(L, Fbn, Fbtt)
Next
XIpb = DrszFF("L Fbn Fbtt", Dry)
End Function

Private Function XIpf(I As DLTDH) As Drs
Dim Dr, Dry(), LTT As Drs, Ix As Dictionary, L&, Inpn$, Ffn$
LTT = DLTT(I, "Inp", "Inpn Ffn").D
For Each Dr In Itr(LTT.Dry)
    L = Dr(0)
    Inpn = Dr(1)
    Ffn = Dr(2)
    PushI Dry, Array(L, Inpn, Ffn, ISfx(Ffn), HasFfn(Ffn))
Next
XIpf = DrszFF("L Inpn Ffn IsFx HasFfn", Dry)
'BrwDrs XIpf: Stop
End Function

Private Function XIpw(A As DLTDH) As Drs
Dim Dr, L&, Dta$, T$, Bexpr$, Dry()
For Each Dr In Itr(DLDta(A, "Tbl.Where").D.Dry)
    L = Dr(0)
    Dta = Dr(1)
    T = T1(Dta)
    Bexpr = RmvT1(Dta)
    PushI Dry, Array(L, T, Bexpr)
Next
XIpw = DrszFF("L T Bexpr", Dry)
End Function
Function FnyAzJn(Jn$) As String()
Dim J
For Each J In SyzSS(Jn)
    PushI FnyAzJn, BefOrAll(J, ":")
Next
End Function
Function FnyBzJn(Jn$) As String()
Dim J
For Each J In SyzSS(Jn)
    PushI FnyBzJn, AftOrAll(J, ":")
Next
End Function

Private Function XIpxHasFx(Ipx As Drs, Ipf As Drs) As Drs
XIpxHasFx = LJnDrs(Ipx, Ipf, "Fxn:Inpn", "Ffn:Fx IsFx HasFfn:HasFx", "HasInp")
End Function

Private Function XIpx1(A As DLTDH) As Drs
'Ret::*Ipx::Drs{L T Fxn Ws Stru}
'BrwDrs A.D: Stop
Dim Dr, L&, Dta$, Dry(), C As Drs, IxT%, IxL%, IxFxnDotWs%, IxStru%, _
Stru$, B$, T$, Ix As Dictionary, Fxn$, Ws$
C = DLTTT(A, "FxTbl", "T FxnDotWs Stru").D
Set Ix = DiczAyIx(C.Fny)
IxT = Ix("T")
IxL = Ix("L")
IxFxnDotWs = Ix("FxnDotWs")
IxStru = Ix("Stru")
For Each Dr In Itr(C.Dry)
    L = Dr(IxL)
    T = Dr(IxT)
    B = Dr(IxFxnDotWs)
    Stru = Dr(IxStru)
    Fxn = BefDotOrAll(B)
    Ws = AftDot(B)
    If Fxn = "" Then Fxn = T
    If Ws = "" Then Ws = "Sheet1"
    If Stru = "" Then Stru = T
    PushI Dry, Array(L, T, Fxn, Ws, Stru)
Next
XIpx1 = DrszFF("L T Fxn Ws Stru", Dry)
'BrwDrs XIpx.D: Stop
End Function

'================================================
Private Function XEsSDup(IpsHead As Drs) As String()

End Function
Private Function XEsSMis(IpsHead As Drs, IpbxStru$()) As String()

End Function
Private Function XEsSExa(IpsHead As Drs, IpbxStru$()) As String()

End Function
Private Function XEsSNoFld(IpsHead As Drs, IpsStru$()) As String()

End Function
Private Function XLsszWhT$(Wh As Drs, T)
'Fm:Wh@Ipw::Drs{L T Bexpr}
Dim O&(), Dr
For Each Dr In Itr(Wh.Dry)
    If Dr(1) = T Then
        Push O, Dr(0)
    End If
Next
XLsszWhT = JnSpc(O)
End Function
Private Function XEwTblExa(Ipw As Drs, Tny$()) As String()

End Function
Private Function XEwTblDup(Ipw As Drs) As String()
'Fm:Wh@Ipw::Drs{L T Bexpr}
Dim OLss$(), OT$(), J%, T, Dr, DupTny$(), Dup, O$()
DupTny = AywDup(StrColzDrs(Ipw, "T"))
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
End Function
Private Function XEwTblMis(Ipw As Drs, Tny$()) As String()
'Fm:Wh@Ipw::Drs{L T Bexpr}
Dim OL&(), OT$(), J%, T, Dr, O$()
For Each Dr In Itr(Ipw.Dry)
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
Private Function XEwBexprEmp(Ipw As Drs) As String()
Dim J%, OL&(), OT$(), O$()
'Fm : Wh@Ipw::Drs{L T Bexpr}
Dim Dr, L&, T$, Bexpr$
For Each Dr In Itr(Ipw.Dry)
    Bexpr = Dr(2)
    If Bexpr = "" Then
        L = Dr(0)
        T = Dr(1)
        PushI OL, L
        PushI OT, T
    End If
Next
'===
For J = 0 To UB(OL)
    PushI O, FmtQQ("L#(?) Tbl(?) has no Bexpr", OL(J), OT(J))
Next
XEwBexprEmp = O
End Function

Private Function DoAddInpIfEr(E$(), InpFilSrc$(), LnkImpSrc$()) As String()
If Si(E) = 0 Then Exit Function
Dim O$(): O = E
PushIAy O, LyzNmLy("InpFilSrc", InpFilSrc, EiBeg1)
PushIAy O, LyzNmLy("LnkImpSrc", LnkImpSrc, EiBeg1)
DoAddInpIfEr = O
End Function
Private Function XIpsFx(Ipx As Drs) As String()
'Ret:*IpsFx::Sy{Stru} ! The struSy used by Fx
XIpsFx = AywDist(StrColzDrs(Ipx, "Stru"))
End Function
Private Function XIps1(Ips As Drs, IpsFx$()) As Drs
'Fm :@Ips::Drs{L Stru F Ty E}
'Fm :@IpsFx::Sy{Stru} ! The Stru used in Fxt
'Ret:*Ips1::Drs{ Ips + IsIpsFx}
Dim Dr, Stru$, IxStru%, ODry()
IxStru = IxzAy(Ips.Fny, "Stru")
For Each Dr In Itr(Ips.Dry)
    Stru = Dr(IxStru)
    PushI Dr, HasEle(IpsFx, Stru)
    PushI ODry, Dr
Next
XIps1 = Drs(AddFF(Ips.Fny, "IsIpsFx"), ODry)
'BrwDrs XIps1: Stop
End Function

Private Function XIps(A As DLTDH) As Drs
'Fm :@LTDH::Drs{L T1 Dta IsHdr}
'Ret:*Ips::Drs{L Stru F Ty E}
Dim B As Drs, Dr, L&, T1$, Dta$, F$, Ty$, E$, Stru$, ODry()
B = DLDtazT1Pfx(A, "Stru.").D
For Each Dr In Itr(B.Dry)
    L = Dr(0)
    T1 = Dr(1)
    Dta = Dr(2)
    Stru = T1
    F = ShfT1(Dta)
    Ty = ShfT1(Dta)
    E = RmvSqBkt(Dta)
    PushI ODry, Array(L, Stru, F, Ty, E)
Next
XIps = DrszFF("L Stru F Ty E", ODry)
'BrwDrs XIps: Stop
End Function
Private Property Get Y_InpFilSrc() As String()
Erase XX
X "Inp"
X " DutyPay C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
X " ZHT0    C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\Pricing report(ForUpload).xls"
X " MB52    C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\2018\MB52 2018-01-30.xls"
X " Uom     C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\sales text.xlsx"
X " GLBal   C:\Users\user\Desktop\MHD\SAPAccessReports\TaxRateAlert\TaxRateAlert\Sample\DutyPrepayGLTot.xlsx"
Y_InpFilSrc = XX
Erase XX
End Property

Private Property Get Y_LnkImpSrc() As String()
Erase XX
X Y_InpFilSrc
X "FbTbl"
X "--  Fbn TblNm.."
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
Erase XX
End Property

