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
Const StruEr3$ = "#StruEr1-Excess."
Const StruEr4$ = "#StruEr1-NoFld."
Const FldEr1$ = "#FldEr1-FldDup."
Const FldEr2$ = "#FldEr2-TyEr."
Const OthEr1$ = "#InpEr1-NoFxAndNoFb. "
Const OthEr2$ = "#OthEr2-SectEr. "
Const WhEr1$ = "#WhEr1-TblDup. "
Const WhEr2$ = "#WhEr2-TblMis. "
Const WhEr3$ = "#WhEr3-BexprEmp. "

Private Function XDiiFx(Dii As Drs) As Drs
XDiiFx = DrswColEqExlEqCol(Dii, "IsFx", True)
End Function
Private Function XDiiExi(Dii As Drs) As Drs
XDiiExi = DrswColEqExlEqCol(Dii, "HasFfn", True)
End Function
Private Function XDiiMis(Dii As Drs) As Drs
XDiiMis = DrswColEqExlEqCol(Dii, "HasFfn", False)
End Function
Private Function XFbInpnAy(Dii As Drs) As String()
XFbInpnAy = StrColzDrswColEqSel(Dii, "IsFx", False, "Inpn")
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
Dim P As DLTDH:           P = DLTDH(LnkImpSrc) ' L T1 Dta IsHdr         !
Dim Dii As Drs:         Dii = XDii(P)          ' L Inpn Ffn IsFx HasFfn !
Dim DiiFx As Drs:     DiiFx = XDiiFx(Dii)      '
Dim DiiExi As Drs:   DiiExi = XDiiExi(Dii)     ' L Inpn Ffn IsFx
Dim DiiMis As Drs:   DiiMis = XDiiMis(Dii)     ' L Inpn Ffn IsFx
Dim FbInpnAy$():   FbInpnAy = XFbInpnAy(Dii)   '                  !
Dim Diw As Drs:         Diw = XDiw(P)
'---------------------
Dim Dix As Drs:           Dix = XDix(P)                                      'L T Fxn Ws Stru ! Inp-Fx which is FxTbl
Dim DixHasFx As Drs: DixHasFx = XDixHasFx(Dix, Dii)                          'L T Fxn Ws Stru Fx IsFx HasFx HasInp ! Add
Dim D1 As Drs:           D1 = DrswColEqExlEqCol(DixHasFx, "IsFx", True)
Dim D2 As Drs:           D2 = DrswColEqExlEqCol(D1, "HasFx", True)
Dim D3 As Drs:           D3 = DrswColEqExlEqCol(D2, "HasInp", True)
Dim DixExiFx As Drs:   DixExiFx = D3                                                   ' L T Fxn Ws Stru Fx !
Dim DActWs As Drs:        DActWs = XDActWs(DixExiFx)                                    ' Fxn Ws !
Dim DixHasWs As Drs:   DixHasWs = AddColzExiB(DixExiFx, DActWs, "Fxn Ws", "HasWs") 'L T Fxn Ws Stru Fx HasFx HasWs !
Dim DixExi As Drs:     DixExi = DrswColEqExlEqCol(DixHasWs, "HasWs", True)         '!DixExi   ::Drs{}
Dim DixMis As Drs:     DixMis = DrswColEqExlEqCol(DixHasWs, "HasWs", False)
'---------------------
Dim FxStru$(): FxStru = XFxStru(Dix)         '!FxStru::Sy{Stru} ! the stru used by Fx
Dim Dis As Drs:   Dis = XDis(P)              '!Dis   ::Drs{L Stru F Ty E}
Dim Dis1 As Drs:  Dis1 = XDis1(Dis, FxStru)   '!Dis1  ::Drs{L Stru F Ty E IsFxStru}
'---------------------
Dim Dib As Drs:  Dib = DLTT(P, "FbTbl", "Fbn Fbtt").D '*Dib::Drs{L Fbn Fbtt}
Dim FxTny$():   FxTny = XFxTny(Dix)
Dim FbTny$():  FbTny = XFbTny(Dib)
Dim Tny$():    Tny = Sy(FxTny, FbTny)
'---
Dim DActWsf As Drs: DActWsf = XDActWsf(DixExi)
Dim Dixf As Drs:    Dixf = XDixf(DixExi, Dis)
Dim DixfMis As Drs: DixfMis = XDixfMis(Dixf, DActWsf)
'---
Dim EiInpnDup$(): EiInpnDup = WEiInpnDup(Dii)
Dim EiFfnDup$(): EiFfnDup = WEiFfnDup(Dii)
Dim EiFfnMis$(): EiFfnMis = WEiFfnMis(DiiMis)
Dim I$(): I = Sy(EiInpnDup, EiFfnDup, EiFfnMis)
'
Dim ExTblDup$(): ExTblDup = WExTblDup(Dix)
Dim ExFxnDup$(): ExFxnDup = WExFxnDup(Dix)
Dim ExFxnMis$(): ExFxnMis = WExFxnMis(Dix, DiiExi)
Dim ExWsMis$(): ExWsMis = WExWsMis(DixMis, DActWs)
Dim ExWsMisFld$(): ExWsMisFld = WExWsMisFld(DixfMis, DActWsf)
Dim ExWsMisFldTy$(): ExWsMisFldTy = WExWsMisFldTy(Dixf, DActWsf)
Dim ExStruMis$(): ExStruMis = WExStruMis()
Dim X$(): X = Sy(ExFxnDup, ExFxnMis, ExWsMis, ExWsMisFld, ExWsMisFldTy)
'
Dim EbFbnDup$(): EbFbnDup = WEbFbnDup(Dib)
Dim EbFbnMis$(): EbFbnMis = WEbFbnMis(Dib, FbInpnAy)
Dim EbTblDup$(): EbTblDup = WEbTblDup(Dib)
Dim EbTblMis$(): EbTblMis = WEbTblMis
Dim EbStruMis$(): EbStruMis = WEbStruMis()
Dim B$(): B = Sy(EbFbnDup, EbFbnMis, EbTblDup, EbTblMis, EbStruMis)
'
Dim EsDup$(): EsDup = WEsDup()
Dim EsMis$(): EsMis = WEsMis()
Dim EsExcess$(): EsExcess = WEsExcess()
Dim EsNoFld$(): EsNoFld = WEsNoFld()
Dim S$(): S = Sy(EsDup, EsMis, EsExcess, EsNoFld)
'
Dim EfFldDup$(): EfFldDup = WEfFldDup()
Dim EfTyEr$(): EfTyEr = WEfTyEr(XDTyEr(Dis1))
Dim F$(): F = Sy(EfFldDup, EfTyEr)
'
Dim EwTblDup$(): EwTblDup = WEwTblDup(Diw)
Dim EwTblMis$(): EwTblMis = WEwTblMis(Diw, Tny)
Dim EwBexprEmp$(): EwBexprEmp = WEwBexprEmp(Diw)
Dim W$(): W = Sy(EwTblDup, EwTblMis, EwBexprEmp)
'
Dim EoNoFxAndNoFb$: EoNoFxAndNoFb = WEoNoFxAndNoFb(Dix, Dib)
Dim EoSectEr$(): EoSectEr = WEoSectEr(P)
Dim O$(): O = Sy(EoNoFxAndNoFb, EoSectEr)

ErzLnk = Sy(I, X, B, S, F, W, O)
End Function
Private Function XDixf(DixExi As Drs, Dis As Drs) As Drs
Dim O As Drs, Dr, E$, IxF%, IE%, J&
O = JnDrs(DixExi, Dis, "Stru", "F Ty E")
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
XDixf = O
'BrwDrs3 O, DixExi, Dis, NN:="Dixf DixExi Dis": Stop
End Function
Private Function XDixfMisFld() As Drs

End Function
Private Function XIsShtTy(ShtTy$, IsFxStru As Boolean) As Boolean
Select Case True
Case Not IsFxStru And ShtTy = "": XIsShtTy = True
Case Else: XIsShtTy = IsShtTy(ShtTy)
End Select
End Function
'==================================================
Private Function XDTyEr(Dis1 As Drs) As Drs
'Fm : @Dis1::Drs{L Stru F Ty E IsFxStru}
'Ret::*DTyEr::Drs{TyEr L Stru F E}
Dim Fny$(), Dr, ITy%, IL%, IStru%, IxF%, IE%, IxIsFxStru%
Dim Ty$, L&, Stru$, F$, E$, IsFxStru As Boolean, ODry()
Fny = Dis1.Fny
IL = IxzAy(Fny, "L")
IStru = IxzAy(Fny, "Stru")
IxF = IxzAy(Fny, "F")
ITy = IxzAy(Fny, "Ty")
IE = IxzAy(Fny, "E")
IxIsFxStru = IxzAy(Fny, "IsFxStru")
For Each Dr In Itr(Dis1.Dry)
    Ty = Dr(ITy)
    IsFxStru = Dr(IxIsFxStru)
    L = Dr(IL)
    Stru = Dr(IStru)
    F = Dr(IxF)
    E = Dr(IE)
    If Not XIsShtTy(Ty, IsFxStru) Then
        PushI ODry, Array(Ty, L, Stru, F, E, IsFxStru)
    End If
Next
XDTyEr = DrszFF("TyEr L Stru F E IsFxStru", ODry)
End Function

Private Function WDActTbl(DiiFb As Drs) As Drs
Dim Dr, T, J%, IFbn$, IFb$, Dry()
For Each Dr In Itr(DiiFb.Dry)
    IFbn = Dr(1)
    IFb = Dr(2)
    For Each T In Tni(Db(IFb))
        PushI Dry, Array(IFbn, T)
    Next
Next
WDActTbl = DrszFF("Fbn T", Dry)
End Function

Private Function XDActWs(DixExiFx As Drs) As Drs
'Fm :DixExi::Drs{L T Fxn Ws Stru Fx}
'Ret:*WsAct::Drs{Fxn Ws}
Dim A As Drs, Dr, Fxn$, Fx$, Wsn, Dry()
A = DrswDist(DixExiFx, "Fxn Fx")
For Each Dr In Itr(A.Dry)
    Fxn = Dr(0)
    Fx = Dr(1)
    For Each Wsn In Wny(Fx)
        PushI Dry, Array(Fxn, Wsn)
    Next
Next
XDActWs = DrszFF("Fxn Ws", Dry)
'BrwDrs XDActWs: Stop
End Function

Function AddFF(Fny$(), FF$) As String()
AddFF = Sy(Fny, SyzSS(FF))
End Function

Private Function WExWsMisFld(DixfMis As Drs, DActWsf As Drs) As String()
'Fm : @DixfMis :: Drs{L Ws Fxn Fx}
'Fm : @DActWsf:: SSAy{Fxn Ws F Ty}
If NoReczDrs(DixfMis) Then Exit Function
Dim OFx$(), OFxn$(), OWs$(), O$(), Fxn, Fx$, Ws$, Mis As Drs, Act As Drs, J%, O1$()
IntoColApzDistDrs DixfMis, "Fxn Fx Ws", OFxn, OFx, OWs
'====
PushI O, "Some columns in ws is missing"
For Each Fxn In OFxn
    Fxn = OFxn(J)
    Fx = OFx(J)
    Ws = OWs(J)
    Mis = DrswCCCEqExlEqCol(DixfMis, "Fxn Fx Ws", Fxn, Fx, Ws)
    Act = DrswCCCEqExlEqCol(DActWsf, "Fxn Fx Ws", Fxn, Fx, Ws)
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
WExWsMisFld = O
End Function
Private Function WExWsMisFldTy(Dixf As Drs, DActWsf As Drs) As String()
'Fm : @DixFld:: Drs{Fxn Ws Stru Dixf Ty Fx} ! Where HasFx and HasWs and Not HasFld
'Fm : @WsActFld::Drs{Fxn Ws Dixf Ty}
'Dim OFxn$(), J%, Fxn$, Fx$, Act$(), Lno&(), Ws$(), DActWsf()
'OFxn = AywDist(StrColzDrs(Dix4, "Fxn"))
''====
'If Si(OFxn) = 0 Then Exit Function
'PushI WExWsMis, "Some expected ws not found"
'For J = 0 To UB(OFxn)
'    Fxn = OFxn(J)
'    Fx = ValzDrswColEqSel(Dix4, "Fxn", Fxn, "Fx")
'    DActWsf = DrswColEqSel(Dix4, "Fxn", Fxn, "L Ws").Dry
'    Lno = LngAyzDryC(DActWsf, 0)
'    Ws = SyzDryC(DActWsf, 1)
'
'    Act = RmvT1zAy(AywT1(WsAct, Fxn)) '*WsActPerFxn::Sy{WsAct}
'    PushIAy WExWsMis, XMisWs_OneFx(Fxn, Fx, Lno, Ws, Act)
'Next
End Function

Private Function WExWsMis(DixMis As Drs, DActWs As Drs) As String()
Dim OFxn$(), J%, Fxn$, Fx$, Act$(), Lno&(), Ws$(), ActWsnn$, DixMisi As Drs, O$()
OFxn = AywDist(StrColzDrs(DixMis, "Fxn"))
'====
If Si(OFxn) = 0 Then Exit Function
PushI O, "Some expected ws not found"
For J = 0 To UB(OFxn)
    Fxn = OFxn(J)
    Fx = ValzDrswColEqSel(DixMis, "Fxn", Fxn, "Fx")
    DixMisi = DrswColEqSel(DixMis, "Fxn", Fxn, "L Ws")
    ActWsnn = TermLin(FstCol(DrswColEqExlEqCol(DActWs, "Fxn", Fxn)))
    '-
    Erase XX
    X "Fxn    : " & Fxn
    X "Fx pth : " & Pth(Fx)
    X "Fx fn  : " & Fn(Fx)
    X "Act ws : " & ActWsnn
    X LyzNmDrs("Mis ws : ", DixMisi)
    PushIAy O, TabAy(XX)
Next
Erase XX
WExWsMis = O
End Function

Private Sub ZZ_ErzLnk()
Brw ErzLnk(Y_LnkImpSrc)
End Sub

Private Function WEiFfnDup(Dii As Drs) As String()
Dim OLss$(), OFfn$(), Dr, L&, Ffn$, Dup, LAy&(), J%, O$()
For Each Dup In Itr(AywDup(StrColzDrs(Dii, "Ffn")))
    LAy = LngAyzDrswColEqSel(Dii, "Ffn", Dup, "L")
    PushI OLss, JnSpc(LAy)
    PushI OFfn, Dup
Next
'=====
If Si(OLss) = 0 Then Exit Function
For J = 0 To UB(OLss)
    PushI O, InpEr2 & FmtQQ("L#(?) Ffn(?) Ffn is duplicated", OLss(J), OFfn(J))
Next
WEiFfnDup = O
End Function

Private Function WEiFfnMis(DiiMis As Drs) As String()
If NoReczDry(DiiMis.Dry) Then Exit Function
Dim O$()
O = LyzNmDrs(InpEr1 & " file missing: ", DiiMis, MaxColWdt:=200)
WEiFfnMis = O
End Function

Private Function WEiInpnDup(Dii As Drs) As String()
'Fm: @Dii::Drs{L Inpn Inpn IsFx}
Dim OLss$(), OInpn$(), Dr, L&, Inpn$, Dup, LAy&(), J%, O$()
For Each Dup In Itr(AywDup(StrColzDrs(Dii, "Inpn")))
    LAy = LngAyzDrswColEqSel(Dii, "Inpn", Dup, "L")
    PushI OLss, JnSpc(LAy)
    PushI OInpn, Dup
Next
'=====
If Si(OLss) = 0 Then Exit Function
For J = 0 To UB(OLss)
    PushI O, InpEr1 & FmtQQ("L#(?) Inpn(?) Inpn is duplicated", OLss(J), OInpn(J))
Next
WEiInpnDup = O
End Function
Private Function WExTblDup(Dix As Drs) As String()

End Function
Private Function WExFxnDup(Dix As Drs) As String()

End Function
Private Function WExFxnMis(Dix As Drs, DiiExi As Drs) As String()

End Function
Private Function WExStruMis() As String()

End Function
Private Function WEbFbnDup(Dib As Drs) As String()
'Fm: Fb@Dib::Drs{L Fbn Fbtt}
Dim Ix As Dictionary, IxL%, IxFbn%, Dup, LAy&(), Fbn, Lss, OLss$(), OFbn$(), J%
Set Ix = DiczAyIx(Dib.Fny)
IxL = Ix("L")
IxFbn = Ix("Fbn")
For Each Dup In Itr(AywDup(DistColzDrs(Dib, "Fbn")))
    LAy = LngAyzDrswColEqSel(Dib, "Fbn", Dup, "L")
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
WEbFbnDup = O
End Function

Private Function WEbFbnMis(Dib As Drs, FbInpnAy$()) As String()
'Fm: @FbInpAy::Sy{Inpn}
'Fm: Fb@Dib::Drs{L Fbn Fbtt}
Dim Ix As Dictionary, IxL%, IxFbn%, Dr, Fbn$, OL&(), L&, OFbn$(), J%, Inpn, O$()
Set Ix = DiczAyIx(Dib.Fny)
IxL = Ix("L")
IxFbn = Ix("Fbn")
For Each Dr In Itr(Dib.Dry)
    Fbn = Dr(IxFbn)
    If Not HasEle(FbInpnAy, Fbn) Then
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
PushI O, vbTab & FmtQQ("Total (?)-Fbn are defined are:", Si(FbInpnAy))
For Each Inpn In Itr(FbInpnAy)
    PushI O, vbTab & vbTab & Inpn
Next
WEbFbnMis = O
End Function

Private Function WEbTblDup(Dib As Drs) As String()
'Fm: Dib@Dib::Drs{L Fbn Fbtt}
Dim J&, OL&(), OFbtt$(), Fbtt$, L&, IxL%, IxFbtt%, Ix As Dictionary, B$, Dr, O$()
Set Ix = DiczAyIx(Dib.Fny)
IxL = Ix("L")
IxFbtt = Ix("Fbtt")
For Each Dr In Itr(Dib.Dry)
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
WEbTblDup = O
End Function

Private Function WEbStruMis() As String()

End Function
Private Function WEbTblMis() As String()

End Function

Private Function WEfFldDup() As String()

End Function

Private Function WEoNoFxAndNoFb$(Dix As Drs, Dib As Drs)
If Si(Dix.Dry) > 0 Then Exit Function
If Si(Dib.Dry) > 0 Then Exit Function
WEoNoFxAndNoFb = OthEr1 & "Both [FxTbl] and [FbTbl] sections are missing"
End Function

Private Function WEfTyEr(DTyEr As Drs) As String()
'Fm:DTyEr@DE?::Drs{ErTy L Stru F E}
If NoReczDrs(DTyEr) Then Exit Function
Dim O$()
PushI O, FldEr2 & "Valid Ty are: ...."
PushIAy O, AddPfxzAy(FmtDrs(DTyEr), vbTab)
WEfTyEr = O
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

Private Function WEoSectEr(A As DLTDH) As String()
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
WEoSectEr = O
End Function

Private Function XDixfMis(Dixf As Drs, DActWsf As Drs) As Drs
'xfMis
'Fm: @Dixf::Drs{}
'Ret:*DixfMis::Drs{}
Dim A As Drs, B As Drs, O As Drs
A = LJnDrs(Dixf, DActWsf, "Fxn Ws E:F", "Ty:ActTy", "HasF")
B = DrswColEqExlEqCol(A, "HasF", False)
O = DrpColzDrs(B, "ActTy")
'BrwDrs4 Dixf, ActWsf, DActWsf, O: Stop
XDixfMis = O
End Function

Private Function XDActWsf(DixExi As Drs) As Drs
'Fm : DixExi@DixExi::Drs{L T Fxn Ws Stru Fx}
'Ret::*DActWsf::Drs{?? Fxn Ws F Ty}
Dim Dr, IDr, O As Drs, OFny$(), ODry(), F$, Ty$, Fx$, Ws$, IFx%, IWs%
'BrwDrs DixExi.D: Stop
IFx = IxzAy(DixExi.Fny, "Fx")
IWs = IxzAy(DixExi.Fny, "Ws")
For Each Dr In Itr(DixExi.Dry)
    Fx = Dr(IFx)
    Ws = Dr(IWs)
    For Each IDr In Itr(DFTyzFxw(Fx, Ws).Dry)
        PushI ODry, AddAy(Dr, IDr)
    Next
Next
OFny = Sy(DixExi.Fny, "F", "Ty")
O = Drs(OFny, ODry)
XDActWsf = O
'BrwDrs2 DixExi, O, NN:="DixExi ActWsf": Stop
End Function
Function XFbTny(Dib As Drs) As String()
'Fm: Dib::Drs{L Fbn Fbtt}
Dim Dr, Fbtt$
For Each Dr In Itr(Dib.Dry)
    Fbtt = Dr(2)
    PushNoDupAy XFbTny, SyzSS(Fbtt)
Next
End Function

Function XFxTny(Dix As Drs) As String()
'Fm: @Dix::Drs{L T Fxn Ws Stru}
Dim Dr
For Each Dr In Itr(Dix.Dry)
    PushNoDup XFxTny, Dr(1)
Next
End Function

Private Function XDib(I As DLTDH) As Drs
Dim Dr, L&, Dta$, Fbn$, Fbtt$, Dry()
For Each Dr In Itr(DLDta(I, "FbTbl").D.Dry)
    L = Dr(0)
    Dta = Dr(1)
    Fbn = T1(Dta)
    Fbtt = RmvT1(Dta)
    PushI Dry, Array(L, Fbn, Fbtt)
Next
XDib = DrszFF("L Fbn Fbtt", Dry)
End Function

Private Function XDii(I As DLTDH) As Drs
Dim Dr, Dry(), LTT As Drs, Ix As Dictionary, L&, Inpn$, Ffn$
LTT = DLTT(I, "Inp", "Inpn Ffn").D
For Each Dr In Itr(LTT.Dry)
    L = Dr(0)
    Inpn = Dr(1)
    Ffn = Dr(2)
    PushI Dry, Array(L, Inpn, Ffn, IsFx(Ffn), HasFfn(Ffn))
Next
XDii = DrszFF("L Inpn Ffn IsFx HasFfn", Dry)
'BrwDrs XDii: Stop
End Function

Private Function XDiw(A As DLTDH) As Drs
Dim Dr, L&, Dta$, T$, Bexpr$, Dry()
For Each Dr In Itr(DLDta(A, "Tbl.Where").D.Dry)
    L = Dr(0)
    Dta = Dr(1)
    T = T1(Dta)
    Bexpr = RmvT1(Dta)
    PushI Dry, Array(L, T, Bexpr)
Next
XDiw = DrszFF("L T Bexpr", Dry)
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

Private Function XDixHasFx(Dix As Drs, Dii As Drs) As Drs
XDixHasFx = LJnDrs(Dix, Dii, "Fxn:Inpn", "Ffn:Fx IsFx HasFfn:HasFx", "HasInp")
End Function

Private Function XDix(A As DLTDH) As Drs
'Ret::*Dix::Drs{L T Fxn Ws Stru}
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
XDix = DrszFF("L T Fxn Ws Stru", Dry)
'BrwDrs XDix.D: Stop
End Function

'================================================
Private Function WEsDup() As String()

End Function
Private Function WEsMis() As String()

End Function
Private Function WEsExcess() As String()

End Function
Private Function WEsNoFld() As String()

End Function
Private Function XLsszWhT$(Wh As Drs, T)
'Fm:Wh@Diw::Drs{L T Bexpr}
Dim O&(), Dr
For Each Dr In Itr(Wh.Dry)
    If Dr(1) = T Then
        Push O, Dr(0)
    End If
Next
XLsszWhT = JnSpc(O)
End Function
Private Function WEwTblDup(Diw As Drs) As String()
'Fm:Wh@Diw::Drs{L T Bexpr}
Dim OLss$(), OT$(), J%, T, Dr, DupTny$(), Dup, O$()
DupTny = AywDup(StrColzDrs(Diw, "T"))
For Each Dup In Itr(DupTny)
    PushI OLss, XLsszWhT(Diw, Dup)
    PushI OT, Dup
Next
'===
If Si(OLss) = 0 Then Exit Function
For J = 0 To UB(OLss)
    PushI O, FmtQQ("L#(?) Tbl(?) Tbl are dup", OLss(J), OT(J))
Next
WEwTblDup = O
End Function
Private Function WEwTblMis(Diw As Drs, Tny$()) As String()
'Fm:Wh@Diw::Drs{L T Bexpr}
Dim OL&(), OT$(), J%, T, Dr, O$()
For Each Dr In Itr(Diw.Dry)
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
WEwTblMis = O
End Function
Private Function WEwBexprEmp(Diw As Drs) As String()
Dim J%, OL&(), OT$(), O$()
'Fm : Wh@Diw::Drs{L T Bexpr}
Dim Dr, L&, T$, Bexpr$
For Each Dr In Itr(Diw.Dry)
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
WEwBexprEmp = O
End Function

Private Function DoAddInpIfEr(E$(), InpFilSrc$(), LnkImpSrc$()) As String()
If Si(E) = 0 Then Exit Function
Dim O$(): O = E
PushIAy O, LyzNmLy("InpFilSrc", InpFilSrc, EiBeg1)
PushIAy O, LyzNmLy("LnkImpSrc", LnkImpSrc, EiBeg1)
DoAddInpIfEr = O
End Function
Private Function XFxStru(Dix As Drs) As String()
'Ret:*FxStru::Sy{Stru} ! The struSy used by Fx
XFxStru = AywDist(StrColzDrs(Dix, "Stru"))
End Function
Private Function XDis1(Dis As Drs, FxStru$()) As Drs
'Fm :@Dis::Drs{L Stru F Ty E}
'Fm :@FxStru::Sy{Stru} ! The Stru used in Fxt
'Ret:*Dis1::Drs{ Dis + IsFxStru}
Dim Dr, Stru$, IxStru%, ODry()
IxStru = IxzAy(Dis.Fny, "Stru")
For Each Dr In Itr(Dis.Dry)
    Stru = Dr(IxStru)
    PushI Dr, HasEle(FxStru, Stru)
    PushI ODry, Dr
Next
XDis1 = Drs(AddFF(Dis.Fny, "IsFxStru"), ODry)
'BrwDrs XDis1: Stop
End Function

Private Function XDis(A As DLTDH) As Drs
'Fm :@LTDH::Drs{L T1 Dta IsHdr}
'Ret:*Dis::Drs{L Stru F Ty E}
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
XDis = DrszFF("L Stru F Ty E", ODry)
'BrwDrs XDis: Stop
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

Sub Z()
ZZ_ErzLnk
End Sub
