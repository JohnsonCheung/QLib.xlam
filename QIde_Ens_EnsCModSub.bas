Attribute VB_Name = "QIde_Ens_EnsCModSub"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_Ens_CSub."

Private Function LinzEptCMod$(Mdn)
LinzEptCMod = FmtQQ("Const CMod$ = ""?.""", Mdn)
End Function

Function IxzCnst&(Src$(), Cnstn)
Dim L, O&
For Each L In Itr(Src)
    If CnstnzL(L) = Cnstn Then IxzCnst = O: Exit Function
    O = O + 1
Next
IxzCnst = -1
End Function
Private Function XCSubLzMthLy$(MthLy$(), Mthn$)
If Not XIsUsingCSub(MthLy) Then Exit Function
XCSubLzMthLy = XCSubLzMthn(Mthn)
End Function

Private Function XCSubLzMthn$(Mthn$)
XCSubLzMthn = FmtQQ("Const CSub$ = CMod & ""?""", Mthn)
End Function

Private Function XCSubLno&(MthLy$(), MthLno&)
Dim I&: I = IxzCnst(MthLy, "CSub")
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

Sub EnsCModSubP(Optional Rpt As EmRpt = EmRpt.EiPushOnly)
EnsCModSubzP CPj, Rpt
End Sub

Sub EnsCModSubM(Optional Rpt As EmRpt)
EnsCModSubzM CMd, Rpt
End Sub

Sub EnsCModSubzP(P As VBProject, Optional Rpt As EmRpt = EmRpt.EiPushOnly)
Dim C As VBComponent
Erase XX
For Each C In P.VBComponents
    EnsCModSubzM C.CodeModule, Rpt
Next
If IsPushzRpt(Rpt) Then Brw XX
End Sub

Private Function XCSubAv(L&, Mthn$, MthLy$())
Dim Cnstn$:   Cnstn = IIf(Mthn = "*Dcl", "CMod", "CSub")
Dim Ix&:         Ix = IxzCnst(MthLy, Cnstn)
Dim ActL$:            If Ix > 0 Then ActL = MthLy(Ix)
Dim Ix1&:             If Ix = -1 Then Ix1 = NxtIxzSrc(MthLy, 0) Else Ix1 = Ix
Dim Lno&:       Lno = L + Ix1
            XCSubAv = Array(ActL, Lno)
'Insp "QIde_Ens_EnsCModSub.XCSubAv", "Inspect", "Oup(XCSubAv) L Mthn MthLy", XCSubAv, L, Mthn, MthLy: Stop
End Function

Private Function XAct(Mth As Drs) As Drs
'Fm Mth : L E Mdy Ty Mthn MthLin MthLy
'Ret    : L Mthn MthLy ActL Lno        ! ActL & Lno: If Ty=*Dcl, they are the CModL & CModLnom Otherwise, CSubL and CSubLno @@
Dim A As Drs: A = SelDrs(Mth, "L Mthn MthLy")
Dim Dr, Dry(): For Each Dr In Itr(A.Dry)
    Dim L&:           L = Dr(0)
    Dim Mthn$:     Mthn = Dr(1)
    Dim MthLy$(): MthLy = Dr(2)
    Dim Av:          Av = XCSubAv(L, Mthn, MthLy) ' ActL Lno ! If Mthn="*Dcl", MthLy will be Dcl,
                                                  '          ! ActL will the CModL from MthLy.  It may fnd or "" if not fnd
                                                  '          ! Lno  will the Lno if fnd or the LnozFstCd-of-MthLy, if not fnd
                                                  '          ! If Mthn<>"*Dcl", Mthn & MthLy will be normal
                                                  '          ! ActL will be the CSubL from MthLy.  If may fnd or "" if not fnd
                                                  '          ! Lno  will the CSubLno if fnd or the NxtLno of MthLin (Using L & MthLy to fnd)

    PushI Dry, AyzAdd(Dr, Av)
Next
XAct = DrszFF("L Mthn MthLy ActL Lno", Dry)
'Insp "QIde_Ens_EnsCModSub.XAct", "Inspect", "Oup(XAct) Mth", FmtDrs(XAct), FmtDrs(Mth): Stop
End Function

Private Function XEpt(Act As Drs) As Drs
'Fm Act : L Mthn MthLy ActL Lno      ! ActL & Lno: If Ty=*Dcl, they are the CModL & CModLnom Otherwise, CSubL and CSubLno
'Ret    : L Mthn MthLy ActL Lno EptL @@
Dim IxMthLy%, IxMthn%: AsgIx Act, "MthLy Mthn", IxMthLy, IxMthn
Dim Dr, Dry(): For Each Dr In Itr(Act.Dry)
    Dim MthLy$(): MthLy = Dr(IxMthLy)
    Dim Mthn$:     Mthn = Dr(IxMthn)
    Dim CSubL$:   CSubL = XCSubLzMthLy(MthLy, Mthn)
    PushI Dr, CSubL
    PushI Dry, Dr
Next
XEpt = AddColzFFDry(Act, "EptL", Dry)
'Insp "QIde_Ens_EnsCModSub.XEpt", "Inspect", "Oup(XEpt) Act", FmtDrs(XEpt), FmtDrs(Act): Stop
End Function

Function EnsCModSubzM(M As CodeModule, Optional Rpt As EmRpt) As Boolean
'-- Prepare Data -------------------------------------------------------------------------------------------------------
Dim Mth As Drs: Mth = DMthc(M)                   ' L E Mdy Ty Mthn MthLin MthLy
Dim Act As Drs: Act = XAct(Mth)                  ' L Mthn MthLy ActL Lno        ! ActL & Lno: If Ty=*Dcl, they are the CModL & CModLnom Otherwise, CSubL and CSubLno
Dim Ept As Drs: Ept = XEpt(Act)                  ' L Mthn MthLy ActL Lno EptL
Dim Dif As Drs: Dif = DrseCeqC(Ept, "ActL EptL") ' L Mthn MthLy ActL Lno EptL   ! Only those Act<>Ept

Dim IsUpd As Boolean: IsUpd = IsUpdzRpt(Rpt)
'== Rpl=================================================================================================================
Dim R1   As Drs:  R1 = ColNe(Dif, "EptL", "")                 ' L Nm MthLy ActL Lno EptL
Dim R2   As Drs:  R2 = ColNe(R1, "ActL", "")                  ' L Nm MthLy ActL Lno EptL
Dim Rpl  As Drs: Rpl = SelDrszAs(R2, "L EptL:NewL ActL:OldL") ' L NewL OldL
Dim ORpl As Drs:       If IsUpd Then RplLin M, Rpl

'== Dlt=================================================================================================================
Dim D1   As Drs:  D1 = ColEq(Dif, "EptL", "")            ' L Nm MthLy ActL Lno EptL
Dim D2   As Drs:  D2 = ColNe(D1, "ActL", "")             ' L Nm MthLy ActL Lno EptL
Dim Dlt  As Drs: Dlt = SelDrszAs(D2, "Mthn L ActL:OldL") ' Mthn L OldL
Dim ODlt As Drs:       If IsUpd Then DltLinzD M, Dlt

'== Ins=================================================================================================================
Dim I1   As Drs:  I1 = ColEq(Dif, "ActL", "")            ' L Nm MthLy ActL Lno EptL
Dim I2   As Drs:  I2 = ColNe(I1, "EptL", "")             ' L Nm MthLy ActL Lno EptL
Dim Ins  As Drs: Ins = SelDrszAs(I2, "Mthn L EptL:NewL") ' Mthn L NewL
Dim OIns As Drs:       If IsUpd Then InsLinzD M, Ins

'== Return True is any Dif
EnsCModSubzM = HasReczDrs(Dif)

Dim IsRpt  As Boolean:  IsRpt = IsRptzRpt(Rpt)
Dim IsPush As Boolean: IsPush = IsPushzRpt(Rpt)
If IsRpt Or IsPush Then
    Dim NCSub%: NCSub = CntColNe(Ept, "EptL", "")
    Dim Msg$(): Msg = XMsg(M, Rpl, Dlt, Ins, NCSub)
    If IsRpt Then Brw Msg
    If IsPush Then X Msg
End If
'Insp CSub, Msg, "Rpl Dlt Ins", FmtDrs(Rpl), FmtDrs(Dlt), FmtDrs(Ins)
End Function

Private Sub XPush(Nm$, Drs As Drs, ONy$(), OAv())
If HasReczDrs(Drs) Then
    PushI ONy, Nm
    PushI OAv, FmtDrs(Drs)
End If
End Sub
Private Function XMsgI$(A As Drs, Nm$)
Dim N%: N = NReczDrs(A)
If N = 0 Then
    XMsgI = "No" & Nm
Else
    XMsgI = "N" & Nm & "(" & N & ")"
End If
End Function
Private Function XMsg(M As CodeModule, Rpl As Drs, Dlt As Drs, Ins As Drs, NCSub%) As String()
Dim MRpl$: MRpl = XMsgI(Rpl, "Rpl")
Dim MIns$: MIns = XMsgI(Ins, "Ins")
Dim MDlt$: MDlt = XMsgI(Dlt, "Dlt")
Dim Msg$:  Msg = FmtQQ("EnsCModCSubzM: ? ? ? NCSub(?) Md(?)", MRpl, MDlt, MIns, NCSub, Mdn(M))
Dim Ny$(), Av()
XPush "Rpl", Rpl, Ny, Av
XPush "Dlt", Dlt, Ny, Av
XPush "Ins", Ins, Ny, Av
XMsg = LyzFunMsgNyAv(CSub, Msg, Ny, Av)
End Function
    
Private Sub ZZZ()
QIde_Ens_CModSub:
End Sub

Sub Z()
EnsCModSubzM Md("QIde_Ens_MthMdy"), EiRptOnly
End Sub
