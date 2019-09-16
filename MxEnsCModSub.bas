Attribute VB_Name = "MxEnsCModSub"
Option Compare Text
Option Explicit
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxEnsCModSub."

Private Function LinzEptCMod$(Mdn)
LinzEptCMod = FmtQQ("Const CMod$ = ""?.""", Mdn)
End Function

Private Function XCSubLzMthLy$(MthLy$(), Mthn$)
If Not XIsUsingCSub(MthLy) Then Exit Function
XCSubLzMthLy = XCSubLzMthn(Mthn)
End Function

Private Function XCSubLzMthn$(Mthn$)
XCSubLzMthn = FmtQQ("Const CSub$ = CMod & ""?""", Mthn)
End Function

Private Function XCSubLno&(MthLy$(), MthLno&)
Dim I&: I = CnstIx(MthLy, "CSub")
If I = 0 Then Exit Function
XCSubLno = I + MthLno
End Function

Private Function XIsUsingCSub(MthLy$()) As Boolean
Dim L
XIsUsingCSub = True
For Each L In Itr(MthLy)
    If HasSubStr(L, "CSub, ") Then Exit Function
    If HasSubStr(L, "(CSub") Then Exit Function
Next
XIsUsingCSub = False
End Function

Sub EnsCModSubP(Optional Upd As EmUpd, Optional Osy)
EnsCModSubzP CPj, Upd, Osy
End Sub

Sub EnsCModSubM(Optional Upd As EmUpd, Optional Osy)
EnsCModSubzM CMd, Upd, Osy
End Sub

Private Sub EnsCModSubzP(P As VBProject, Optional Upd As EmUpd, Optional Osy)
Dim C As VBComponent
For Each C In P.VBComponents
    EnsCModSubzM C.CodeModule, Upd
Next
End Sub

Private Function XCSubAv(L&, Mthn$, MthLy$())
Dim Cnstn$:   Cnstn = IIf(Mthn = "*Dcl", "CMod", "CSub")
Dim Ix&:         Ix = CnstIx(MthLy, Cnstn)
Dim ActL$:            If Ix > 0 Then ActL = MthLy(Ix)
Dim Ix1&:             If Ix = -1 Then Ix1 = NxtIxzSrc(MthLy, 0) Else Ix1 = Ix
Dim Lno&:       Lno = L + Ix1
            XCSubAv = Array(ActL, Lno)
'Insp "QIde_Ens_EnsCModSub.XCSubAv", "Inspect", "Oup(XCSubAv) L Mthn MthLy", XCSubAv, L, Mthn, MthLy: Stop
End Function

Private Function XDoAct(Mth As Drs) As Drs
'Fm Mth : L E Mdy Ty Mthn MthLin MthLy
'Ret    : L Mthn MthLy ActL Lno        ! ActL & Lno: If Ty=*Dcl, they are the CModL & CModLnom Otherwise, CSubL and CSubLno @@
Dim A As Drs: A = SelDrs(Mth, "L Mthn MthLy")
Dim Dr, Dy(): For Each Dr In Itr(A.Dy)
    Dim L&:           L = Dr(0)
    Dim Mthn$:     Mthn = Dr(1)
    Dim MthLy$(): MthLy = Dr(2)
    Dim Av:          Av = XCSubAv(L, Mthn, MthLy) ' ActL Lno ! If Mthn="*Dcl", MthLy will be Dcl,
                                                  '          ! ActL will the CModL from MthLy.  It may fnd or "" if not fnd
                                                  '          ! Lno  will the Lno if fnd or the LnozFstCd-of-MthLy, if not fnd
                                                  '          ! If Mthn<>"*Dcl", Mthn & MthLy will be normal
                                                  '          ! ActL will be the CSubL from MthLy.  If may fnd or "" if not fnd
                                                  '          ! Lno  will the CSubLno if fnd or the NxtLno of MthLin (Using L & MthLy to fnd)

    PushI Dy, AddAy(Dr, Av)
Next
XDoAct = DrszFF("L Mthn MthLy ActL Lno", Dy)
'Insp "QIde_Ens_EnsCModSub.XDoAct", "Inspect", "Oup(XDoAct) Mth", FmtCellDrs(XDoAct), FmtCellDrs(Mth): Stop
End Function

Private Function XDoEpt(Act As Drs) As Drs
'Fm Act : L Mthn MthLy ActL Lno      ! ActL & Lno: If Ty=*Dcl, they are the CModL & CModLnom Otherwise, CSubL and CSubLno
'Ret    : L Mthn MthLy ActL Lno EptL @@
Dim IxMthLy%, IxMthn%: AsgIx Act, "MthLy Mthn", IxMthLy, IxMthn
Dim Dr, Dy(): For Each Dr In Itr(Act.Dy)
    Dim MthLy$(): MthLy = Dr(IxMthLy)
    Dim Mthn$:     Mthn = Dr(IxMthn)
    Dim CSubL$:   CSubL = XCSubLzMthLy(MthLy, Mthn)
    PushI Dr, CSubL
    PushI Dy, Dr
Next
XDoEpt = AddColzFFDy(Act, "EptL", Dy)
'Insp "QIde_Ens_EnsCModSub.XDoEpt", "Inspect", "Oup(XDoEpt) Act", FmtCellDrs(XDoEpt), FmtCellDrs(Act): Stop
End Function
Private Sub Z_EnsCModSubzM()
Dim M As CodeModule: Set M = Md("QIde_Ens_EnsCModSub")
EnsCModSubzM M
End Sub
Private Sub EnsCModSubzM(M As CodeModule, Optional Upd As EmUpd, Optional Osy)

'-- Prepare Data -------------------------------------------------------------------------------------------------------
Dim DoMth As Drs: DoMth = DoMthczM(M)                    ' L E Mdy Ty Mthn MthLin Mthl
Stop
Dim DoAct As Drs: DoAct = XDoAct(DoMth)              ' L Mthn MthLy ActL Lno        ! ActL & Lno: If Ty=*Dcl, they are the CModL & CModLnom Otherwise, CSubL and CSubLno
Dim DoEpt As Drs: DoEpt = XDoEpt(DoAct)              ' L Mthn MthLy ActL Lno EptL
Dim DoDif As Drs: DoDif = DeCeqC(DoEpt, "ActL EptL") ' L Mthn MthLy ActL Lno EptL   ! Only those Act<>Ept

'== Rpl=================================================================================================================
Dim XUpd As Boolean: XUpd = IsUpd(Upd)
Dim Rpl As Drs: Rpl = XRpl(DoDif)          '
                      RplLin M, Rpl      ' <==
Dim Dlt As Drs: Dlt = XDlt(DoDif)
                      DltLinzD M, Dlt    ' <==
Dim Ins As Drs: Ins = XIns(DoDif)
                      InsLinzD M, Ins    ' <==

'== Return True is any Dif==============================================================================================
If IsRpt(Upd, Osy) Then
    Dim Msg$(): Msg = XMsg(M, DoEpt, Rpl, Dlt, Ins)
    RptMsg Upd, Msg
    Osy = AddOsy(Osy, Msg)
End If
'Insp CSub, Msg, "Rpl Dlt Ins", FmtCellDrs(Rpl), FmtCellDrs(Dlt), FmtCellDrs(Ins)
End Sub

Sub RptMsg(Upd As EmUpd, Msg$())
If IsRptU(Upd) Then
    Dmp Msg
End If
End Sub

Private Function XIns(Dif As Drs) As Drs

End Function
Private Function XDlt(Dif As Drs) As Drs

End Function
Private Function XRpl(Dif As Drs) As Drs

End Function

Private Sub XPush(NM$, Drs As Drs, ONy$(), OAv())
If HasReczDrs(Drs) Then
    PushI ONy, NM
    PushI OAv, FmtCellDrs(Drs)
End If
End Sub

Private Function XMsgI$(A As Drs, NM$)
Dim N%: N = NReczDrs(A)
If N = 0 Then
    XMsgI = "No" & NM
Else
    XMsgI = "N" & NM & "(" & N & ")"
End If
End Function

Private Function XMsg(M As CodeModule, Ept As Drs, Rpl As Drs, Dlt As Drs, Ins As Drs) As String()
Dim NCSub%: NCSub = CntColNe(Ept, "EptL", "")
Dim MRpl$: MRpl = XMsgI(Rpl, "Rpl")
Dim MIns$: MIns = XMsgI(Ins, "Ins")
Dim MDlt$: MDlt = XMsgI(Dlt, "Dlt")
Dim Msg$:  Msg = FmtQQ("EnsCModCSubzM: ? ? ? NCSub(?) Md(?)", MRpl, MDlt, MIns, NCSub, Mdn(M))
Dim Ny$(), Av()
XPush "Rpl", Rpl, Ny, Av
XPush "Dlt", Dlt, Ny, Av
XPush "Ins", Ins, Ny, Av
XMsg = LyzFunMsgNyAv(CSub, Msg, Ny, Av)
'Insp "QIde_Ens_EnsCModSub.XMsg", "Inspect", "Oup(XMsg) M Rpl Dlt Ins NCSub", XMsg, Mdn(M), FmtCellDrs(Rpl), FmtCellDrs(Dlt), FmtCellDrs(Ins), NCSub: Stop
End Function
    
Private Sub Z()
EnsCModSubP
End Sub