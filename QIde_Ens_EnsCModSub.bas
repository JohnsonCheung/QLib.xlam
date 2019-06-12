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
Private Function W_CSubLzMthLy$(MthLy$(), Mthn$)
If Not W_IsUsingCSub(MthLy) Then Exit Function
W_CSubLzMthLy = W_CSubLzMthn(Mthn)
End Function

Private Function W_CSubLzMthn$(Mthn$)
W_CSubLzMthn = FmtQQ("Const CSub$ = CMod & ""?""", Mthn)
End Function

Private Function W_CSubLno&(MthLy$(), MthLno&)
Dim I&: I = IxzCnst(MthLy, "CSub")
If I = 0 Then Exit Function
W_CSubLno = I + MthLno
End Function

Private Function W_IsUsingCSub(MthLy$()) As Boolean
Dim L
W_IsUsingCSub = True
For Each L In Itr(MthLy)
    If HasSubStr(L, "CSub, ") Then Exit Function
    If HasSubStr(L, "(CSub") Then Exit Function
Next
W_IsUsingCSub = False
End Function

Sub EnsCModSubP(Optional Rpt As EmRpt = EmRpt.EiPushOnly)
EnsCModSubzP CPj, Rpt
End Sub

Sub EnsCModSubM()
EnsCModSubzM CMd
End Sub

Sub EnsCModSubzP(P As VBProject, Optional Rpt As EmRpt = EmRpt.EiPushOnly)
Dim C As VBComponent
Erase XX
For Each C In P.VBComponents
    EnsCModSubzM C.CodeModule, Rpt
Next
If Rpt = EiPushOnly Then
    Brw XX
    Stop
End If
End Sub

Private Function XA1_Av(L&, Mthn$, MthLy$()) As Variant()
'Ret Av : ActL Lno ! If Mthn="*Dcl", MthLy will be Dcl,
'                             !    ActL will the CModL from MthLy.  It may fnd or "" if not fnd
'                             !    Lno  will the Lno if fnd or the LnozFstCd-of-MthLy, if not fnd
'                             ! If Mthn<>"*Dcl", Mthn & MthLy will be normal
'                             !    ActL will be the CSubL from MthLy.  If may fnd or "" if not fnd
'                             !    Lno  will the CSubLno if fnd or the NxtLno of MthLin (Using L & MthLy to fnd) @@
Dim Cnstn$: Cnstn = IIf(Mthn = "*Dcl", "CMod", "CSub")
Dim Ix&:       Ix = IxzCnst(MthLy, Cnstn)
Dim ActL$: If Ix > 0 Then ActL = MthLy(Ix)
If Ix = -1 Then Ix = NxtIxzSrc(MthLy, 0)
Dim Lno&:  Lno = L + Ix
XA1_Av = Array(ActL, Lno)
End Function

Private Function XA_Act(Mth As Drs) As Drs
'Fm  Mth :  ' L E Mdy Ty Mthn MthLin MthLy
'Ret Act :  ' L Mthn MthLy ActL Lno        !  ! ActL & Lno: If Ty=*Dcl, they are the CModL & CModLnom Otherwise, CSubL and CSubLno @@
Const CmPfx$ = "XA1_"
Dim A As Drs: A = SelDrs(Mth, "L Mthn MthLy")
Dim Dr, Dry(): For Each Dr In Itr(A.Dry)
    Dim L&, Mthn$, X: AsgAp Dr, L, Mthn, X
    Dim MthLy$(): MthLy = X
    Dim Av(): Av = XA1_Av(L, Mthn, MthLy) ' ActL Lno ! If Mthn="*Dcl", MthLy will be Dcl,
'                             !    ActL will the CModL from MthLy.  It may fnd or "" if not fnd
'                             !    Lno  will the Lno if fnd or the LnozFstCd-of-MthLy, if not fnd
'                             ! If Mthn<>"*Dcl", Mthn & MthLy will be normal
'                             !    ActL will be the CSubL from MthLy.  If may fnd or "" if not fnd
'                             !    Lno  will the CSubLno if fnd or the NxtLno of MthLin (Using L & MthLy to fnd)

    Dim ActL$, Lno&: AsgAp Av, ActL, Lno
    PushI Dry, Array(L, Mthn, MthLy, ActL, Lno)
Next
XA_Act = DrszFF("L Mthn MthLy ActL Lno", Dry)
'Insp "QIde_Ens_CModSub.EnsCModSubzM", "Inspect", "Oup(XA_Act) Act Mth",FmtDrs(Act), FmtDrs(Act), FmtDrs(Mth): Stop
End Function

Private Function XA_Ept(Act As Drs) As Drs
'Fm  Act :  ' L Mthn MthLy ActL Lno      ! ActL & Lno: If Ty=*Dcl, they are the CModL & CModLnom Otherwise, CSubL and CSubLno
'Ret Ept :  ' L Mthn MthLy ActL Lno EptL11 @@
Dim IxMthLy%, IxMthn%: AsgIx Act, "MthLy Mthn", IxMthLy, IxMthn
Dim Dr, Dry(): For Each Dr In Itr(Act.Dry)
    Dim MthLy$(): MthLy = Dr(IxMthLy)
    Dim Mthn$: Mthn = Dr(IxMthn)
    Dim CSubL$: CSubL = W_CSubLzMthLy(MthLy, Mthn)
    PushI Dr, CSubL
    PushI Dry, Dr
Next
XA_Ept = AddColzFFDry(Act, "EptL", Dry)
'Insp "QIde_Ens_CModSub.EnsCModSubzM", "Inspect", "Oup(XA_Ept) Ept Act",FmtDrs(Ept), FmtDrs(Ept), FmtDrs(Act): Stop
End Function
Function EnsCModSubzM(M As CodeModule, Optional Rpt As EmRpt) As Boolean
'-- Prepare Data -------------------------------------------------------------------------------------------------------
Const CmPfx$ = "XA_"
Dim IsUpd As Boolean: IsUpd = IsUpdzRpt(Rpt)

Dim Mth As Drs: Mth = DMthEL(M)   ' L E Mdy Ty Mthn MthLin MthLy
Dim Act As Drs: Act = XA_Act(Mth) ' L Mthn MthLy ActL Lno        ! ActL & Lno: If Ty=*Dcl, they are the CModL & CModLnom Otherwise, CSubL and CSubLno
Dim Ept As Drs: Ept = XA_Ept(Act) ' L Mthn MthLy ActL Lno EptL
Dim Dif As Drs: Dif = DrseCeqC(Ept, "ActL EptL") ' L Mthn MthLy ActL Lno EptL ! Only those Act<>Ept
'== Rpl=================================================================================================================
Dim R1   As Drs:  R1 = ColNe(Dif, "EptL", "")              ' L Nm MthLy ActL Lno EptL
Dim R2   As Drs:  R2 = ColNe(R1, "ActL", "")               ' L Nm MthLy ActL Lno EptL
Dim Rpl  As Drs: Rpl = SelDrszAs(R2, "L EptL:NewL ActL:OldL") ' L NewL OldL
Dim ORpl As Drs:       If IsUpd Then RplLin M, Rpl

'== Dlt=================================================================================================================
Dim D1   As Drs:  D1 = ColEq(Dif, "EptL", "")         ' L Nm MthLy ActL Lno EptL
Dim D2   As Drs:  D2 = ColNe(D1, "ActL", "")          ' L Nm MthLy ActL Lno EptL
Dim Dlt  As Drs: Dlt = SelDrszAs(D2, "Mthn L ActL:OldL") ' Mthn L OldL
Dim ODlt As Drs:       If IsUpd Then DltLinzD M, Dlt

'== Ins=================================================================================================================
Dim I1   As Drs:  I1 = ColEq(Dif, "ActL", "")         ' L Nm MthLy ActL Lno EptL
Dim I2   As Drs:  I2 = ColNe(I1, "EptL", "")          ' L Nm MthLy ActL Lno EptL
Dim Ins  As Drs: Ins = SelDrszAs(I2, "Mthn L EptL:NewL") ' Mthn L NewL
Dim OIns As Drs:       If IsUpd Then InsLinzD M, Ins

EnsCModSubzM = HasReczDrs(Dif)

Dim IsRpt As Boolean:   IsRpt = IsRptzRpt(Rpt)
Dim IsPush As Boolean: IsPush = IsPushzRpt(Rpt)
If IsRpt Or IsPush Then
    Dim Msg$:  Msg = "EnsCModCSubzM: " & Mdn(M)
    Dim L1$(): L1 = FmtDrs(Rpl)
    Dim L2$(): L2 = FmtDrs(Dlt)
    Dim L3$(): L3 = FmtDrs(Ins)
    Dim S$():   S = LyzFunMsgNap(CSub, Msg, "Rpl Dlt Ins", L1, L2, L3)
    If IsRpt Then Brw S
    If IsPush Then X S
'    Insp CSub, Msg, "Rpl Dlt Ins", FmtDrs(Rpl), FmtDrs(Dlt), FmtDrs(Ins)
End If
End Function

Private Sub ZZZ()
QIde_Ens_CModSub:
End Sub

Sub Z()
EnsCModSubzM Md("QIde_Ens_MthMdy"), EiRptOnly
End Sub
