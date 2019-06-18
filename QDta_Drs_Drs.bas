Attribute VB_Name = "QDta_Drs_Drs"
Option Compare Text
Option Explicit
Const Asm$ = "QDta"
Const Ns$ = "Dta.Ds"
Private Const CMod$ = "BDrs."
Type DrSepr
    DtaSep As String
    DtaQuote As String
    LinSep As String
    LinQuote As String
End Type
Enum EmTblFmt
    EiTblFmt = 0
    EiSSFmt = 1
End Enum
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
Private Type GpDrs
    GpDrs As Drs
    RLvlGpIx() As Long
End Type
    
Function Drsz4TRstLy(T4RstLy$(), FF$) As Drs
Dim I, Dry(): For Each I In Itr(T4RstLy)
    PushI Dry, Syz4TRst(T4RstLy)
Next
Drsz4TRstLy = DrszFF(FF, Dry)
End Function
Function DrszTRstLy(TRstLy$(), FF$) As Drs
Dim I, Dry(): For Each I In Itr(TRstLy)
    PushI Dry, SyzTRst(I)
Next
DrszTRstLy = DrszFF(FF, Dry)
End Function
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

Function HasReczDry(Dry()) As Boolean
HasReczDry = Si(Dry) > 0
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
Function LasDr(A As Drs)
LasDr = LasEle(A.Dry)
End Function

Sub XXX()
Dim A$: A = "123456789"
Mid(A, 2, 3) = "abcdex"
Debug.Print A
End Sub

Function FmtLNewO(L_NewL_OldL As Drs, Org_L_Lin As Drs) As String()
'Fm  : L_NewL_OldL ! Assume all NewL and OldL are nonEmp and <>
'Ret : LinesAy !
Dim SDry(): SDry = DryzSel(Org_L_Lin.Dry, LngAp(0, 1))
Dim S As Drs: S = DrszFF("L Lin", SDry)
Dim D As Drs: D = DrseCeqC(L_NewL_OldL, "NewL OldL")
Dim NewL As Drs: NewL = LDrszJn(S, D, "L", "NewL")
Dim Gpno As Drs: Gpno = FmtLNewOGpno(NewL)
Dim NLin As Drs: NLin = FmtLNewONLin(Gpno)
Dim Lines As Drs: Lines = FmtLNewOLines(NLin)
Dim OneG As Drs: OneG = FmtNewOneG(NLin)
FmtLNewO = StrColzDrs(OneG, "Lines")
End Function
Private Function FmtLNewOLines(NLin As Drs) As Drs
'Fm NLin: L Gpno NLin SNewL
'Ret Lines: L Gpno Lines
Dim Dr, L&, Gpno&, Lines$, NLin_$, SNewL
Dim Dry()
'Insp SNewL should have some Emp
'    Erase Dry
'    For Each Dr In NLin.Dry
'        PushI Dr, IsEmpty(Dr(2))
'        PushI Dry, Dr
'    Next
'    BrwDrs DrszFF("L Gpno NLin SNewL Emp", Dry)
'    Erase Dry
For Each Dr In Itr(NLin.Dry)
    AsgAp Dr, L, Gpno, NLin_, SNewL
    If IsEmpty(SNewL) Then
        Lines = NLin_
    Else
        Lines = NLin_ & vbCrLf & SNewL
    End If
    PushI Dry, Array(L, Gpno, Lines)
Next
FmtLNewOLines = DrszFF("L Gpno Lines", Dry)
'BrwDrs FmtLNewOLines: Stop
End Function
Private Function FmtNewOneG(NLin As Drs) As Drs
'Fm  D: L Gpno NLin SNewL !
'Ret E: Gpno Lines ! Gpno now become uniq
Dim O$(), L&, LasG&, Dr, Dry(), Gpno&, NLin_$, SNewL
If NoReczDrs(NLin) Then Exit Function
LasG = NLin.Dry(0)(1)
For Each Dr In Itr(NLin.Dry)
    AsgAp Dr, L, Gpno, NLin_, SNewL
    If LasG <> Gpno Then
        PushI Dry, Array(Gpno, JnCrLf(O))
        Erase O
        LasG = Gpno
    End If
    PushI O, NLin_
    If Not IsEmpty(SNewL) Then PushI O, SNewL
Next
If Si(O) > 0 Then PushI Dry, Array(Gpno, JnCrLf(O))
FmtNewOneG = DrszFF("Gpno Lines", Dry)
End Function

Private Function FmtLNewOGpno(NewL As Drs) As Drs
'Fm  NewL: L Lin NewL ! NewL may empty, when non-Emp, NewL <> Lin
'Ret D: L Lin NewL Gpno ! Gpno is running from 1:
'                      !   all conseq Lin with Emp-NewL is one group
'                      !   each non-Emp-NewL is one gp
Dim IGpno&, Dr, Dry(), Lin, NewL_, LasEmp As Boolean, Emp As Boolean

'For Each Dr In Itr(NewL.Dry)
'    PushI Dr, IsEmpty(Dr(2))
'    PushI Dry, Dr
'Next
'BrwDry Dry
'Erase Dry
'Stop
LasEmp = True
IGpno = 0
For Each Dr In Itr(NewL.Dry)
    Lin = Dr(1)
    NewL_ = Dr(2)
    Emp = IsEmpty(NewL_)
    If Not Emp Then If Lin = NewL_ Then Stop
    If IsEmpty(Lin) Then Stop
    Select Case True
    Case Not Emp: IGpno = IGpno + 1
    Case Emp And Not LasEmp: IGpno = IGpno + 1
    Case Else
    End Select
    PushI Dr, IGpno
    PushI Dry, Dr
    LasEmp = Emp
Next
FmtLNewOGpno = DrszFF("L Lin NewL Gpno", Dry)
End Function
Private Function FmtLNewONLin(Gpno As Drs) As Drs
'Fm  Gpno: L Lin NewL Gpno
'Ret E: L Gpno NLin SNewL ! NLin=L# is in front; SNewL = Spc is in front, only when nonEmp
Dim MaxL&: MaxL = MaxzAy(LngAyzDrs(Gpno, "L"))
Dim NDig%: NDig = Len(CStr(MaxL))
Dim S$: S = Space(NDig + 1)
Dim Dry(), Dr, L&, Lin$, NewL, IGpno&, NLin$, SNewL
For Each Dr In Itr(Gpno.Dry)
    AsgAp Dr, L, Lin, NewL, IGpno
    NLin = AlignR(L, NDig) & " " & Lin
    If IsEmpty(NewL) Then
        SNewL = Empty
    Else
        SNewL = S & NewL
    End If
    PushI Dry, Array(L, IGpno, NLin, SNewL)
Next
FmtLNewONLin = DrszFF("L Gpno NLin SNewL", Dry)
End Function
Function EmpLNewO() As Drs
EmpLNewO = LNewO(EmpAv())
End Function
Function LNewO(Dry()) As Drs
LNewO = DrszFF("L NewL OldL", Dry)
End Function

Function IxzDryDr&(Dry(), Dr)
Dim IDr, O&: For Each IDr In Itr(Dry)
    If IsEqAy(IDr, Dr) Then IxzDryDr = O: Exit Function
    O = O + 1
Next
IxzDryDr = -1
End Function
Sub Z_Agr()
BrwDrs Agr(DMthP, "Mdn Ty", "Mthn")
End Sub

Function Agr(D As Drs, Gpcc$, Optional C) As Drs
'Fm  D : ..{Gpcc} {C}.. ! it has columns-Gpcc and column-C
'Ret   : {Gpcc} {C}     ! where C is group of column-C @@
If C = "" Then
    Agr = SelDist(D, Gpcc)
    Exit Function
End If
Dim FF$: FF = Gpcc & " " & C

Dim OCol(), OKey()
    Dim A As Drs: A = DrszSel(D, FF)
    Dim Dr: For Each Dr In Itr(A.Dry)
        Dim V: V = Pop(Dr)
        Dim Ix&: Ix = IxzDryDr(OKey, Dr)
        If Ix = -1 Then
            PushI OKey, Dr
            PushI OCol, Array(V)
        Else
            PushI OCol(Ix), V
        End If
    Next
Dim ODry()
    Dim J&: For J = 0 To UB(OCol)
        PushI OKey(J), OCol(J)
        PushI ODry, OKey(J)
    Next
Agr = DrszFF(FF, ODry)
End Function
Private Function AgrCntzDry(Dry()) As Variant()
Dim Dr: For Each Dr In Itr(Dry)
    PushI Dr, Si(Pop(Dr))
    PushI AgrCntzDry, Dr
Next
End Function
Private Sub Z_AgrCnt()
BrwDrs AgrCnt(DMthP, "Mdn")
End Sub
Function AgrCnt(D As Drs, Gpcc$) As Drs
Dim A As Drs: A = Agr(D, Gpcc)
Dim Dry(): Dry = AgrCntzDry(A.Dry)
AgrCnt = Drs(D.Fny, Dry)
End Function

Function AgrMin(D As Drs, Gpcc$, MinC$) As Drs
Dim Dry()
    Dim A As Drs: A = Agr(D, Gpcc, MinC)
    Dim Dr: For Each Dr In Itr(A.Dry)
        Dim Col(): Col = Pop(Dr)
        PushI Dr, MinzAy(Col)
        PushI Dry, Dr
    Next
AgrMin = Drs(D.Fny, Dry)
End Function

Function AgrMax(D As Drs, Gpcc$, MaxC$) As Drs
Dim Dry()
    Dim A As Drs: A = Agr(D, Gpcc, MaxC)
    Dim Dr: For Each Dr In Itr(A.Dry)
        Dim Col(): Col = Pop(Dr)
        PushI Dr, MaxzAy(Col)
        PushI Dry, Dr
    Next
AgrMax = Drs(D.Fny, Dry)
End Function

Function DiczRenFF(RenFF$) As Dictionary
Set DiczRenFF = New Dictionary
Dim Ay$(): Ay = SyzSS(RenFF)
Dim V: For Each V In SyzSS(RenFF)
    If HasSubStr(V, ":") Then
        DiczRenFF.Add Bef(V, ":"), Aft(V, ":")
    Else
        Thw CSub, "Invalid RenFF.  all Sterm has have [:]", "RenFF", RenFF
    End If
Next
End Function

Function FnyzRen(Fny$(), RenFF$) As String()
Dim D As Dictionary: Set D = DiczRenFF(RenFF)
Dim F: For Each F In Fny
    If D.Exists(F) Then
        PushI FnyzRen, D(F)
    Else
        PushI FnyzRen, F
    End If
Next
End Function
Function DrszRen(D As Drs, RenFF$) As Drs
DrszRen = Drs(FnyzRen(D.Fny, RenFF), D.Dry)
End Function

Function DrszSplitSS(D As Drs, SSCol$) As Drs
'Fm D     : It has a col @SSCol
'Fm SSCol : It is a col nm in @D whose value is SS.
'Ret  : a drs of sam ret but more rec by split@SSCol col to multi record
Dim I%: I = IxzAy(D.Fny, SSCol)
Dim Dr, Dry(): For Each Dr In Itr(D.Dry)
    Dim S: For Each S In Itr(SyzSS(Dr(I)))
        Dr(I) = S
        PushI Dry, Dr
    Next
Next
DrszSplitSS = Drs(D.Fny, Dry)
End Function

Function RLvlGpIx(Dry()) As Long()

End Function

Sub Z_GpCol()
Dim Col():            Col = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
Dim RLvlGpIx&(): RLvlGpIx = LngAp(1, 1, 1, 3, 3, 2, 2, 3, 0, 0)
Dim G():                G = ColGp(Col, RLvlGpIx)
Stop
End Sub

Private Function ColGp(Col(), RLvlGpIx&()) As Variant()
'Fm Col      : Col to gp
'Fm RLvlGpIx : Each V in Col is mapped to GpIx by this RLvlGpix @@
ThwIf_DifSi Col, RLvlGpIx, CSub
Dim MaxGpIx&: MaxGpIx = MaxzAy(RLvlGpIx)
Dim O(): ReDim O(MaxGpIx)
Dim I&: For I = 0 To MaxGpIx
    O(I) = Array()
Next
I = 0
Dim V: For Each V In Itr(Col)
    Dim GpIx&: GpIx = RLvlGpIx(I)
    PushI O(GpIx), V
    I = I + 1
Next
ColGp = O
End Function

Private Sub Z_AgrWdt()
BrwDrs AgrWdt(DMthP, "Mdn Ty", "Mthn")
End Sub

Private Function AgrWdt(D As Drs, Gpcc$, C) As Drs
Dim A As Drs: A = Agr(D, Gpcc, C)
Dim Dr, Dry(): For Each Dr In Itr(A.Dry)
    Dim Col(): Col = Pop(Dr)
    PushI Dr, WdtzAy(Col)
    PushI Dry, Dr
Next
Dim Fny$(): Fny(UB(Fny)) = "W" & C
AgrWdt = Drs(Fny, Dry)
End Function
Private Function DrszAliC__Dry(WiWdtDry(), Cix%) As Variant()

End Function
Function DryGp(Dry(), Cxy&()) As Variant()
'Fm Dry : Dry to be gp.  It has all col as stated in @Cxy.
'Fm Cxy : Gpg which col of @Dry
'Ret    : Ay-of-Dry.  Each ele is a subset of @Dry in same gp.  @@
Dim Key(): Key = DryzSel(Dry, Cxy) ' sel the key only
Dim RGp(): RGp = RxyGp(Key)        ' gp the key into :RxyGp
Dim Rxy: For Each Rxy In Itr(RGp)          ' each @Rxy is a gp
    Dim DGp()                      ' putting each gp of @Dry in @DGp
    Dim R: For Each R In Rxy
        PushI DGp, Dry(R)          ' pushing @Dry-rec-R to @DGp
    Next
    PushI DryGp, DGp               ' pushing whole gp of @Dry rec (in @DGp) as one ele to @DryGp (the ret) '<===
    Erase DGp
Next
End Function

Function RxyGp(Dry()) As Variant()
'Fm Dry : to be gp.
'Ret    : Gp of Rxy.  Each gp will have sam val of rec.
Dim I&, Key(), O(), Dr, R&: For Each Dr In Itr(Dry)
    I = IxzDryDr(Key, Dr)
    If I = -1 Then
        PushI Key, Dr
        PushI O, LngAp(R)
    Else
        PushI O(I), R
    End If
    R = R + 1
Next
RxyGp = O
End Function

Private Sub DryzAli__AliCol(ODry(), C)
'Fm ODry : the col @C will be aligned
'Fm C    : the column ix
'Ret     : column-@C of @ODry will be aligned
Dim Col(): Col = ColzDry(ODry, C)
Dim ACol$(): ACol = SyzAlign(Col)
Dim J&: For J = 0 To UB(ODry)
    ODry(J)(C) = ACol(J)
Next
End Sub

Private Function DryzAli(Dry(), Cix&()) As Variant()
Dim O(): O = Dry
Dim C: For Each C In Cix
    DryzAli__AliCol O, C
Next
DryzAli = O
End Function
Function DrszAli(D As Drs, Gpcc$, CC$) As Drs
'Fm D : ..@Gpcc..@CC.. ! It has a str col @CC to be alignL and @Gpcc to be gp
'Ret  : @D             ! all col @CC are aligned within the gp and rec will gp together.
If NoReczDrs(D) Then DrszAli = D: Exit Function
Dim Cix&(): Cix = IxyzCC(D, CC)       ' The grouping col ix ay
Dim DGp(): DGp = DryGp(D.Dry, Cix)          ' Grouping the-data (@D.Dry) into DryGp (@DGp) according to @Cix
Dim Dry, ADry(), ODry(): For Each Dry In DGp  'For each data-gp (@Dry)
    ADry = DryzAli(CvAv(Dry), Cix)  ' aligning it (@Dry) into (@ADry)
    PushIAy ODry, ADry              ' pushing the rslt (@ADry-aligned-dry) to oup (@ODry)
Next
DrszAli = Drs(D.Fny, ODry)
End Function

Sub CrtTzAlignCC(D As Database, T$, Fm$, Gpcc$, CC$)
Dim GFny$():       GFny = SyzSS(Gpcc)
Dim CFny$():       CFny = SyzSS(CC)
Dim N&:               N = Si(CFny)
Dim SeqN%():       SeqN = IntSeq(N, 1)               ' Seq
Dim WFny$():       WFny = SyzQAy("#W?", SeqN)    ' Wdt of each CC
Dim WFnySq$():   WFnySq = SyzQteSq(WFny)
Dim MFny$():       MFny = SyzQAy("#M?", SeqN)     ' Max of Wdt
Dim FF$:             FF = Gpcc & " " & CC
Dim WMFny$():     WMFny = SyzAdd(WFny, MFny)
Dim WMFnySq$(): WMFnySq = SyzAdd(WFnySq, SyzQteSq(MFny))

'-- Bld Q* all Sql Stmt --
Dim QO$:           QO = SqlSelzIntoCpy(T, Fm)  ' O :

Dim ColAy$():   ColAy = SyzAyS(WMFnySq, " Integer")
Dim QOAddWM$: QOAddWM = SqlAddColzAy(T, ColAy)     ' O : Gpcc CC [#W?] [#M?]  ! Add col [#W?] [#M?] all are integer

Dim Ey$():         Ey = SyzQAy("Len(?)", CFny)
Dim QOUpdW$:   QOUpdW = SqlUpdzEy(T, WFny, Ey)   ' O     :                  ! Update Wcc as Len(CC)

Dim Eq0$():       Eq0 = SyzAyS(WFnySq, " = 0")
Dim Bexp$:       Bexp = JnAnd(Eq0)
Dim QONoW0$:   QONoW0 = SqlDlt(T, Bexp)                                       ' O     :                  ! Rmv rec with all rmv = 0

Dim TG$:           TG = TmpNm("TGp_")
Dim StrQ$:       StrQ = "Max(x.[#W?]) as [#M?]"
Dim Max$():       Max = SyzQAy(StrQ, SeqN)
Dim SelMax$:   SelMax = JnCommaSpc(SyzAdd(GFny, Max))
Dim Gp$:           Gp = JnCommaSpc(GFny)
Dim QG$:           QG = SqlSelzIntoFmX(TG, SelMax, T, Gp)       ' G : Gpcc Mcc         ! Mcc is Max-Wcc gp by Gpcc
Dim QOUpdM$:   QOUpdM = SqlUpdzJn(T, TG, GFny, MFny)            ' O :

Dim AliEy$():   AliEy = SyzMacro("[{F}] & Space([#M{N}]-[#W{N}]) & '.'", CFny, SeqN)
Dim QOAli$:     QOAli = SqlUpdzEy(T, CFny, AliEy)                            ' O :                  ! align CC

Dim QODrpWM$: QODrpWM = SqlDrpFld(T, WMFny)                                            ' O M : Sam as T         ! Use T.Gpcc and T3.CC

D.Execute QO
D.Execute QOAddWM
D.Execute QOUpdW
D.Execute QONoW0
D.Execute QG
D.Execute QOUpdM
D.Execute QOAli
D.Execute QODrpWM
DrpT D, TG
End Sub

Sub UpdTczFillLas(D As Database, T, F$)
With RszTF(D, T, F)
    Dim Fst As Boolean: Fst = True
    Dim L
    While Not .EOF
        If Fst Then
            Fst = False
              L = .Fields(0).Value
        Else
            If Trim(Nz(.Fields(0).Value, "")) = "" Then
                .Edit
                .Fields(0).Value = L
                .Update
            Else
                L = .Fields(0).Value
            End If
        End If
        .MoveNext
    Wend
    .Close
End With
End Sub
Function DrszFillLasIfB(D As Drs, C$) As Drs
'Fm D : It has a str col C
'Ret  : Fill in the blank col-C val by las val
Dim LasV$
Dim Fst As Boolean: Fst = True
Dim Ix%: Ix = IxzAy(D.Fny, C)
Dim Dr, Dry(): For Each Dr In Itr(D.Dry)
    Dim V$: V = Dr(Ix)
    If Fst Then
        LasV = V
        Fst = False
    End If
    If V = "" Then
        Dr(Ix) = LasV
    Else
        LasV = V
    End If
    PushI Dry, Dr
Next
DrszFillLasIfB = Drs(D.Fny, Dry)
End Function

Function Drs(Fny$(), Dry()) As Drs
With Drs
    .Fny = Fny
    .Dry = Dry
End With
End Function
Function DrszNewDry(A As Drs, NewDry()) As Drs
With DrszNewDry
    .Fny = A.Fny
    .Dry = NewDry
End With
End Function

Function DrsAddCol(A As Drs, ColNm$, CnstBrk) As Drs
DrsAddCol = Drs(CvSy(AyzAddItm(A.Fny, ColNm)), DryAddColzC(A.Dry, CnstBrk))
End Function

Function EnsColTyzInt(A As Drs, C) As Drs
If NoReczDrs(A) Then EnsColTyzInt = A: Exit Function
Dim O As Drs, J&, Ix%, Dr
Ix = IxzAy(A.Fny, C)
O = A
If IsSy(O.Dry(0)) Then Stop
For Each Dr In Itr(O.Dry)
    Dr(Ix) = CInt(Dr(Ix))
    O.Dry(J) = Dr
    J = J + 1
Next
EnsColTyzInt = O
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
DrswDist = DrszFF(CC, DrywDist(DrszSel(A, CC).Dry))
End Function
Sub AsgCol(A As Drs, CC$, ParamArray OColAp())
Dim OColAv(), J%, Col, C$()
OColAv = OColAp
C = SyzSS(CC)
For J = 0 To UB(OColAv)
    Col = IntozDrsC(OColAv(J), A, C(J))
    OColAp(J) = Col
Next
End Sub

Sub AsgColDist(A As Drs, CC$, ParamArray OColAp())
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

Sub DmpDrs(A As Drs, Optional MaxColWdt% = 100, Optional BrkColnn$, Optional Fmt As EmTblFmt = EiTblFmt)
DmpAy FmtDrs(A, MaxColWdt, BrkColnn$)
End Sub

Function DrpColzFny(A As Drs, Fny$()) As Drs
DrpColzFny = DrszSelCC(A, MinusAy(A.Fny, Fny))
End Function

Function DrszSelCC(A As Drs, CC$) As Drs
Const CSub$ = CMod & "DrszSelCC"
Dim OFny$(): OFny = TermAy(CC)
If Not IsAySub(A.Fny, OFny) Then Thw CSub, "Given FF has some field not in Drs.Fny", "CC Drs.Fny", CC, A.Fny
Dim ODry()
    Dim IAy&()
    IAy = Ixy(A.Fny, OFny)
    ODry = DryzSelColIxy(A.Dry, IAy)
DrszSelCC = Drs(OFny, ODry)
End Function
Function DryzSelColIxy(Dry(), Ixy&()) As Variant()
Dim Dr
For Each Dr In Itr(Dry)
    PushI DryzSelColIxy, AywIxy(Dr, Ixy)
Next
End Function

Function DrsInsCV(A As Drs, C$, V) As Drs
DrsInsCV = Drs(CvSy(AyInsEle(A.Fny, C)), InsColzDryzV(A.Dry, V, IxzAy(A.Fny, C)))
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
Dry = InsColzDryzV(A.Dry, V, Ix)
DrsInsCVIsAftFld = Drs(Fny1, Dry)
End Function
Function IsNeFF(A As Drs, FF$) As Boolean
IsNeFF = JnSpc(A.Fny) <> FF
End Function
Function IsEqDrs(A As Drs, B As Drs) As Boolean
Select Case True
Case Not IsEqAy(A.Fny, B.Fny), Not IsEqDry(A.Dry, B.Dry)
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
Function IsSamDrEleCntzDry(Dry()) As Boolean
If Si(Dry) = 0 Then IsSamDrEleCntzDry = True: Exit Function
Dim C%: C = Si(Dry(0))
Dim Dr
For Each Dr In Itr(Dry)
    If Si(Dr) <> C Then Exit Function
Next
IsSamDrEleCntzDry = True
End Function
Function IsSamDrEleCnt(A As Drs) As Boolean
IsSamDrEleCnt = IsSamDrEleCntzDry(A.Dry)
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
If NoReczDrs(A) Then Exit Function
Dim Nc&, NR&, Dry(), Fny$()
    Fny = A.Fny
    Dry = A.Dry
    Nc = Max(NColzDry(Dry), Si(Fny))
    NR = Si(Dry)
Dim O()
ReDim O(1 To 1 + NR, 1 To Nc)
Dim C&, R&, Dr
    For C = 1 To Si(Fny)
        O(1, C) = Fny(C - 1)
    Next
    For R = 1 To NR
        Dr = Dry(R - 1)
        For C = 1 To Min(Si(Dr), Nc)
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

Function IsEmpDrs(A As Drs) As Boolean
If HasReczDrs(A) Then Exit Function
If Si(A.Fny) > 0 Then Exit Function
IsEmpDrs = True
End Function

Function DrszAdd3(A As Drs, B As Drs, C As Drs) As Drs
Dim O As Drs: O = DrszAdd(A, B)
          DrszAdd3 = DrszAdd(O, C)
End Function

Function DrszAdd(A As Drs, B As Drs) As Drs
If IsEmpDrs(A) Then DrszAdd = B: Exit Function
If IsEmpDrs(B) Then DrszAdd = A: Exit Function
If Not IsEqAy(A.Fny, B.Fny) Then Thw CSub, "Dif Fny: Cannot add", "A-Fny B-Fny", A.Fny, B.Fny
DrszAdd = Drs(A.Fny, CvAv(AyzAdd(A.Dry, B.Dry)))
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

Private Sub Z_GpDicDKG()
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

Private Sub Z_CntDiczDrs()
Dim Drs As Drs, Dic As Dictionary
'Drs = Vbe_Mth12Drs(CVbe)
Set Dic = CntDiczDrs(Drs, "Nm")
BrwDic Dic
End Sub

Private Sub Z_DrszSel()
BrwDrs DrszSel(SampDrs1, "A B D")
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


Function AddColz2(A As Drs, FF$, C1, C2) As Drs
Dim Fny$(), Dry()
Fny = AyzAdd(A.Fny, TermAy(FF))
Dry = DryAddColzCC(A.Dry, C1, C2)
AddColz2 = Drs(Fny, Dry)
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
BKeyDry = DryzSel(B.Dry, IB)
For Each Dr In Itr(A.Dry)
    KA = AywIxy(Dr, IA)
    If HasDr(BKeyDry, KA) Then
        PushI Dr, True
    Else
        PushI Dr, False
    End If
    PushI ODry, Dr
Next
AddColzExiB = Drs(FnyzAddFF(A.Fny, ExiB_FldNm), ODry)
End Function


Function DrszMapAy(Ay, MapFunNN$, Optional FF$) As Drs
Dim Dry(), V: For Each V In Ay
    Dim Dr(): Dr = Array(V)
    Dim F: For Each F In Itr(SyzSS(MapFunNN))
        PushI Dr, Run(F, V)
    Next
    PushI Dry, Dr
Next
Dim A$: A = DftStr(FF, "V " & MapFunNN)
DrszMapAy = DrszFF(A, Dry)
End Function

Function AddColzLen(D As Drs, AsCol$) As Drs
'Fm AsCol : If no as, {Col}Len will be used
'Ret      : add a len col at end using LenCol @@
Dim C$:       C = BefOrAll(AsCol, ":")
Dim LenC$: LenC = AftOrAll(AsCol, ":")
                  If LenC = C Then LenC = C & "Len"
Dim Ix&: Ix = IxzAy(D.Fny, C)
Dim Dry(), Dr: For Each Dr In Itr(D.Dry)
    Dim L%: L = Len(Dr(Ix))
    PushI Dr, L
    PushI Dry, Dr
Next
AddColzLen = DrszAddFF(D, LenC, Dry)
End Function

Function DrszAddCV(A As Drs, C$, V) As Drs
Dim Dr, Dry()
For Each Dr In Itr(A.Dry)
    PushI Dr, V
    PushI Dry, Dr
Next
DrszAddCV = DrszAddFF(A, C, Dry)
End Function

Function DrszAddFF(A As Drs, FF$, NewDry()) As Drs
DrszAddFF = Drs(FnyzAddFF(A.Fny, FF), NewDry)
End Function

Function DrszInsFF(A As Drs, FF$, NewDry()) As Drs
DrszInsFF = Drs(SyzAdd(SyzSS(FF), A.Fny), NewDry)
End Function

Function DrszAddFiller(A As Drs, CC$) As Drs
Dim O As Drs: O = A
Dim C
For Each C In SyzSS(CC)
    O = DrszAddFillerC(O, C)
Next
DrszAddFiller = O
End Function

Private Function DrszAddFillerC(A As Drs, C) As Drs
'Fm   A : ..{C}.. ! @A should have col @C
'Fm   C : #Coln.
'Ret    : a new drs with addition col @F where F = "F" & C and value eq Len-of-Col-@C
If NoReczDrs(A) Then Stop
Dim W%: W = WdtzAy(StrColzDrs(A, C))
Dim I%: I = IxzAy(A.Fny, C)
Dim ODry(): ODry = A.Dry
Dim Dr, J&
For Each Dr In Itr(ODry)
    PushI Dr, W - Len(Dr(I))
    ODry(J) = Dr
    J = J + 1
Next
DrszAddFillerC = Drs(FnyzAddFF(A.Fny, "F" & C), ODry)
End Function

