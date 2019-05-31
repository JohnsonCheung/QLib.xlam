Attribute VB_Name = "QDta_Drs_Drs"
Option Compare Text
Option Explicit
Const Asm$ = "QDta"
Const NS$ = "Dta.Ds"
Private Const CMod$ = "BDrs."
Type DrSepr
    DtaSep As String
    DtaQuote As String
    LinSep As String
    LinQuote As String
End Type
Enum EmTblFmt: EiTblFmt: EiSSFmt: End Enum
Type Drs: Fny() As String: Dry() As Variant: End Type
Type Drss: N As Long: Ay() As Drs: End Type
Enum EmCnt
    EiCntAll
    EiCntDup
    EiCntSng
End Enum
Enum EmCntSrtOpt
    eNoSrt
    eSrtByCnt
    eSrtByItm
End Enum

Function DrSepr(DtaSep$, DtaQuote$, LinSep$, LinQuote$) As DrSepr
With DrSepr
    .DtaQuote = DtaQuote
    .DtaSep = DtaSep
    .LinQuote = LinQuote
    .LinSep = LinSep
End With
End Function

Function DrSeprzEmTblFmt(A As EmTblFmt) As DrSepr
Dim DS$, DQ$, LS$, LQ$
Select Case True
   Case A = EiTblFmt: DQ = "| * |": DS = " | ": LQ = "|-*-|": LS = "-|-"
   Case A = EiSSFmt: DS = " ": LS = " "
End Select
DrSeprzEmTblFmt = DrSepr(DS, DQ, LS, LQ)
End Function

Function DrszF(FF$) As Drs
DrszF = DrszFF(FF, EmpAv)
End Function
Function DrszFF(FF$, Dry()) As Drs
DrszFF = Drs(TermAy(FF), Dry)
End Function
Function HasReczDrs(A As Drs) As Boolean
HasReczDrs = Si(A.Dry) > 0
End Function
Function NoReczDry(Dry()) As Boolean
NoReczDry = Si(Dry) = 0
End Function
Function NoReczDrs(A As Drs) As Boolean
NoReczDrs = NoReczDry(A.Dry)
End Function
Function ValzDrsC(A As Drs, C)
If Si(A.Dry) = 0 Then Thw CSub, "No Rec", "Drs.Fny", A.Fny
ValzDrsC = A.Dry(0)(IxzAy(A.Fny, C))
End Function
Function LasRec(A As Drs) As Drs
If Si(A.Dry) = 0 Then Thw CSub, "No LasRec", "Drs.Fny", A.Fny
LasRec = Drs(A.Fny, Av((LasEle(A.Dry))))
End Function
Function DrszSSAy(SSAy$(), FF$) As Drs
DrszSSAy = DrszFF(FF, DryzSSAy(SSAy))
End Function
Function DryzSSAy(SSAy$()) As Variant()
Dim SS
For Each SS In Itr(SSAy)
    PushI DryzSSAy, TermAy(SS)
Next
End Function
Function IxdzDrs(A As Drs) As Dictionary
Set IxdzDrs = DiczAyIx(A.Fny)
End Function
Function Drs(Fny$(), Dry()) As Drs
With Drs
    .Fny = Fny
    .Dry = Dry
End With
End Function

Function DrsAddCol(A As Drs, ColNm$, CnstBrk) As Drs
DrsAddCol = Drs(CvSy(AddAyItm(A.Fny, ColNm)), DryAddColzC(A.Dry, CnstBrk))
End Function

Function DrsAddIxCol(A As Drs, IxCol As EmIxCol) As Drs
Dim J&, Fny$(), Dry(), I, Dr
Select Case True
Case IxCol = EiNoIx: DrsAddIxCol = A: Exit Function
Case IxCol = EiBeg0
Case IxCol = EiBeg1: J = 1
End Select
Fny = AyInsEle(A.Fny, "Ix")
For Each Dr In Itr(A.Dry)
    Push Dry, AyInsEle(Dr, J)
    J = J + 1
Next
DrsAddIxCol = Drs(Fny, Dry)
End Function

Function AvDrsC(A As Drs, C) As Variant()
AvDrsC = IntozDrsC(Array(), A, C)
End Function
Function DrswDist(A As Drs, CC$) As Drs
DrswDist = DrszFF(CC, DrywDist(SelDrs(A, CC).Dry))
End Function
Sub IntoColApzDrs(A As Drs, CC$, ParamArray OColAp())
Dim OColAv(), J%, Col, C$()
OColAv = OColAp
C = SyzSS(CC)
For J = 0 To UB(OColAv)
    Col = IntozDrsC(OColAv(J), A, C(J))
    OColAp(J) = Col
Next
End Sub

Sub IntoColApzDistDrs(A As Drs, CC$, ParamArray OColAp())
Dim OColAv(), J%, Col, B As Drs, C$()
B = DrswDist(A, CC)
OColAv = OColAp
C = SyzSS(CC)
For J = 0 To UB(OColAv)
    Col = IntozDrsC(OColAv(J), B, C(J))
    OColAp(J) = Col
Next
End Sub

Function IntozDrsC(Into, A As Drs, C)
Dim O, Ix%, Dry(), Dr
Ix = IxzAy(A.Fny, C): If Ix = -1 Then Stop
O = Into
Erase O
Dry = A.Dry
If Si(Dry) = 0 Then IntozDrsC = O: Exit Function
For Each Dr In Dry
    Push O, Dr(Ix)
Next
IntozDrsC = O
End Function
Function DmpRec(A As Drs)
D FmtRec(A)
End Function
Function FmtRec(A As Drs) As String()
Dim Fny$(), Dr, N&, Ix&
Fny = AlignLzAy(A.Fny)
For Each Dr In Itr(A.Dry)
    PushIAy FmtRec, FmtRec_FmAlignedFny_AndDr(Fny, Dr, Ix, N)
    Ix = Ix + 1
Next
End Function
Function IxOfUStr$(Ix&, U&)

End Function

Function FmtReczFnyDr(Fny$(), Dr, Optional Ix& = -1, Optional N& = -1)
FmtReczFnyDr = FmtRec_FmAlignedFny_AndDr(AlignLzAy(Fny), Dr, Ix, N)
End Function

Function FmtRec_FmAlignedFny_AndDr(AlignedFny$(), Dr, Optional Ix& = -1, Optional U& = -1) As String()
PushNonBlank FmtRec_FmAlignedFny_AndDr, IxOfUStr(Ix, U)
End Function

Sub DmpDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColnn$, Optional Fmt As EmTblFmt = EiTblFmt)
DmpAy FmtDrs(A, MaxColWdt, BrkColnn$)
End Sub

Function DrpCny(A As Drs, CC$) As Drs
DrpCny = SelDrsCC(A, MinusAy(A.Fny, Ny(CC)))
End Function

Function SelDrsCC(A As Drs, CC$) As Drs
Const CSub$ = CMod & "SelDrsCC"
Dim OFny$(): OFny = TermAy(CC)
If Not IsAySub(A.Fny, OFny) Then Thw CSub, "Given FF has some field not in Drs.Fny", "CC Drs.Fny", CC, A.Fny
Dim ODry()
    Dim IAy&()
    IAy = Ixy(A.Fny, OFny)
    ODry = SelDryColIxy(A.Dry, IAy)
SelDrsCC = Drs(OFny, ODry)
End Function
Function SelDryColIxy(Dry(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI SelDryColIxy, AywIxy(Dr, Ixy)
Next
End Function

Function DrsInsCV(A As Drs, C$, V) As Drs
DrsInsCV = Drs(CvSy(AyInsEle(A.Fny, C)), DryInsColzV(A.Dry, V, IxzAy(A.Fny, C)))
End Function

Function DrsInsCVAft(A As Drs, C$, V, AftFldNm$) As Drs
DrsInsCVAft = DrsInsCVIsAftFld(A, C, V, True, AftFldNm)
End Function

Function DrsInsCVBef(A As Drs, C$, V, BefFldNm$) As Drs
DrsInsCVBef = DrsInsCVIsAftFld(A, C, V, False, BefFldNm)
End Function

Private Function DrsInsCVIsAftFld(A As Drs, C$, V, IsAft As Boolean, FldNm$) As Drs
Dim Fny$(), Dry(), Ix&, Fny1$()
Fny = A.Fny
Ix = IxzAy(Fny, C): If Ix = -1 Then Stop
If IsAft Then
    Ix = Ix + 1
End If
Fny1 = AyInsEle(Fny, FldNm, CLng(Ix))
Dry = DryInsColzV(A.Dry, V, Ix)
DrsInsCVIsAftFld = Drs(Fny1, Dry)
End Function

Function IsEqDrs(A As Drs, B As Drs) As Boolean
Select Case True
Case IsEqAy(A.Fny, B.Fny), IsEqDry(A.Dry, B.Dry)
Case Else: IsEqDrs = True
End Select
End Function

Sub BrwCnt(Ay, Optional Opt As EmCnt)
Brw FmtCntDic(CntDic(Ay, Opt))
End Sub
Function DicItmWdt%(A As Dictionary)
Dim I, O%
For Each I In A.Items
    O = Max(Len(I), O)
Next
DicItmWdt = O
End Function
Private Function CntLyzCntDic(CntDic As Dictionary, CntWdt%) As String()
Dim K
For Each K In CntDic.Keys
    PushI CntLyzCntDic, AlignR(CntDic(K), CntWdt) & " " & K
Next
End Function
Function CntLy(Ay, Optional Opt As EmCnt, Optional SrtOpt As EmCntSrtOpt, Optional IsDesc As Boolean) As String()
Dim D As Dictionary: Set D = CntDic(Ay, Opt)
Dim K
Dim W%: W = DicItmWdt(D)
Dim O$()
Select Case SrtOpt
Case eNoSrt
    CntLy = CntLyzCntDic(D, W)
Case eSrtByCnt
    CntLy = QSrt1(CntLyzCntDic(D, W), IsDesc)
Case eSrtByItm
    CntLy = CntLyzCntDic(SrtDic(D, IsDesc), W)
Case Else
    Thw CSub, "Invalid SrtOpt", "SrtOpt", SrtOpt
End Select
End Function

Function NColzDrs%(A As Drs)
NColzDrs = Max(Si(A.Fny), NColzDry(A.Dry))
End Function

Function NReczDrs&(A As Drs)
NReczDrs = Si(A.Dry)
End Function

Function DrwIxy(Dr(), Ixy&())
Dim U&: U = MaxEle(Ixy)
Dim O: O = Dr
If UB(O) < U Then
    ReDim Preserve O(U)
End If
DrwIxy = AywIxy(O, Ixy)
End Function
Function SelCol(Dry(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI SelCol, AywIxy(Dr, Ixy)
Next
End Function
Function ReOrdCol(A As Drs, BySubFF$) As Drs
Dim SubFny$(): SubFny = TermAy(BySubFF)
Dim OFny$(): OFny = AyReOrd(A.Fny, SubFny)
Dim IAy&(): IAy = Ixy(A.Fny, OFny)
Dim ODry(): ODry = SelCol(A.Dry, IAy)
ReOrdCol = Drs(OFny, ODry)
End Function

Function NRowzColEv&(A As Drs, ColNm$, EqVal)
NRowzColEv = NRowzInDryzColEv(A.Dry, IxzAy(A.Fny, ColNm), EqVal)
End Function

Function SqzDrs(A As Drs) As Variant()
Dim NC&, NR&, Dry(), Fny$()
    Fny = A.Fny
    Dry = A.Dry
    NC = Max(NColzDry(Dry), Si(Fny))
    NR = Si(Dry)
Dim O()
ReDim O(1 To 1 + NR, 1 To NC)
Dim C&, R&, Dr
    For C = 1 To Si(Fny)
        O(1, C) = Fny(C - 1)
    Next
    For R = 1 To NR
        Dr = Dry(R - 1)
        For C = 1 To Min(Si(Dr), NC)
            O(R + 1, C) = Dr(C - 1)
        Next
    Next
SqzDrs = O
End Function

Sub ColApzDrs(A As Drs, CC$, ParamArray OColAp())
Dim Av(): Av = OColAp
Dim C$(): C = SyzSS(CC)
Dim J%, O
For J = 0 To UB(Av)
    O = OColAp(J)
    O = IntozDrsC(O, A, C(J)) 'Must put into O first!!
                              'This will die: OColAp(J) = IntozDrsC(O, A, C(J))
    OColAp(J) = O
Next
End Sub

Function SyzDryC(Dry(), C) As String()
SyzDryC = IntozDryC(EmpSy, Dry, C)
End Function

Function SyzDrsC(A As Drs, ColNm$) As String()
SyzDrsC = IntozDrsC(EmpSy, A, ColNm)
End Function
Function AddDrs(A As Drs, B As Drs) As Drs
If Not IsEqAy(A.Fny, B.Fny) Then Thw CSub, "Dif Fny: Cannot add", "A-Fny B-Fny", A.Fny, B.Fny
AddDrs = Drs(A.Fny, CvAv(AddAy(A.Dry, B.Dry)))
End Function
Sub PushDrs(O As Drss, M As Drs)
With O
    ReDim Preserve .Ay(.N)
    .Ay(.N) = M
    .N = .N + 1
End With
End Sub
Private Function IxyWiNegzSupSubAy(SupAy, SubAy) As Long()
If Not IsSuperAy(SupAy, SubAy) Then Thw CSub, "SupAy & SubAy error", "SupAy SubAy", SupAy, SubAy
Dim J%
For J = 0 To UB(SupAy)
    PushI IxyWiNegzSupSubAy, IxzAy(SubAy, SupAy(J))
Next
End Function
Private Function SelDr(Dr(), IxyWiNeg&()) As Variant()
Dim Ix, U%: U = UB(IxyWiNeg)
For Each Ix In IxyWiNeg
    If IsBet(Ix, 0, U) Then
        PushI SelDr, Dr(Ix)
    Else
        PushI SelDr, Empty
    End If
Next
End Function
Sub ApdDrsSub(O As Drs, M As Drs)
Dim Ixy&(): Ixy = IxyWiNegzSupSubAy(O.Fny, M.Fny)
Dim ODry(): ODry = O.Dry
Dim Dr
For Each Dr In Itr(M.Dry)
    PushI ODry, SelDr(CvAv(Dr), Ixy)
Next
O.Dry = ODry
End Sub
Sub ApdDrs(O As Drs, M As Drs)
If Not IsEqAy(O.Fny, M.Fny) Then Thw CSub, "Fny are dif", "O.Fny M.Fny", O.Fny, M.Fny
Dim UO&, UM&, U&, J&
UO = UB(O.Dry)
UM = UB(M.Dry)
U = UO + UM + 1
ReDim Preserve O.Dry(U)
For J = UO + 1 To U
    O.Dry(J) = M.Dry(J - UO - 1)
Next
End Sub

Private Sub ZZ_GpDicDKG()
Dim Act As Dictionary, Dry(), Dr1, Dr2, Dr3
Dr1 = Array("A", , 1)
Dr2 = Array("A", , 2)
Dr3 = Array("B", , 3)
Dry = Array(Dr1, Dr2, Dr3)
Set Act = DryGpDic(Dry, IntAy(0), 2)
Ass Act.Count = 2
Ass IsEqAy(Act("A"), Array(1, 2))
Ass IsEqAy(Act("B"), Array(3))
Stop
End Sub

Private Sub ZZ_CntDiczDrs()
Dim Drs As Drs, Dic As Dictionary
'Drs = Vbe_Mth12Drs(CVbe)
Set Dic = CntDiczDrs(Drs, "Nm")
BrwDic Dic
End Sub

Private Sub ZZ_SelDrs()
BrwDrs SelDrs(SampDrs1, "A B D")
End Sub

Private Property Get Z_FmtDrs()
GoTo ZZ
ZZ:
DmpAy FmtDrs(SampDrs1)
End Property

Private Sub ZZ()
Dim A As Variant
Dim B()
Dim C As Drs
Dim D$
Dim E%
Dim F$()
DrsAddCol C, D, A
DrsAddCol C, D, A
AddColzValIdzCntzDrs C, D, D
DtzDrs C, D
DrsInsCV C, D, A
End Sub


Function DrsAddCC(A As Drs, FF$, C1, C2) As Drs
Dim Fny$(), Dry()
Fny = AddAy(A.Fny, TermAy(FF))
Dry = DryAddColzCC(A.Dry, C1, C2)
DrsAddCC = Drs(Fny, Dry)
End Function

Function IxyzJnA(Fny$(), Jn$) As Long()
IxyzJnA = IxyzSubAy(Fny, FnyAzJn(Jn))
End Function
Function IxyzJnB(Fny$(), Jn$) As Long()
IxyzJnB = IxyzSubAy(Fny, FnyBzJn(Jn))
End Function
Function AddColzExiB(A As Drs, B As Drs, Jn$, ExiB_FldNm$) As Drs
Dim IA&(), IB&(), Dr, KA(), BKeyDry(), ODry()
IA = IxyzJnA(A.Fny, Jn)
IB = IxyzJnB(B.Fny, Jn)
BKeyDry = SelDry(B.Dry, IB)
For Each Dr In Itr(A.Dry)
    KA = AywIxy(Dr, IA)
    If HasDr(BKeyDry, KA) Then
        PushI Dr, True
    Else
        PushI Dr, False
    End If
    PushI ODry, Dr
Next
AddColzExiB = Drs(AddFF(A.Fny, ExiB_FldNm), ODry)
End Function

