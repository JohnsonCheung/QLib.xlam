Attribute VB_Name = "QTp_SqTp_Sqy"
Option Explicit
Option Compare Text
Private Enum EmStmtTy
    EiDrpStmt
    EiSelStmt
    EiUpdStmt
End Enum
Private Const CMod$ = "MTp_SqyRslt."
Private Const Asm$ = "QTp"
Const U_Into$ = "INTO"
Const U_Sel$ = "SEL"
Const U_SelDis$ = "SELECT DISTINCT"
Const U_Fm$ = "FM"
Const U_Gp$ = "GP"
Const U_Wh$ = "WH"
Const U_And$ = "AND"
Const U_Jn$ = "JN"
Const U_LeftJn$ = "LEFT JOIN"

Const SqBlkTyNN$ = "ER PM SW SQ RM"
Enum EmSwLinOp
    EiOpAnd
    EiOpOr
    EiOpEq
    EiOpNe
End Enum
Type Sw
    StmtSw As Dictionary
    FldSw As Dictionary
End Type
Type SwLin
    Nm As String
    Op As EmSwLinOp
    T1 As String
    T2 As String
    TermAy() As String
End Type
Private Type SwLins: N As Long: Ay() As SwLin: End Type
Private Type EvlSwLinsRslt
    HasEvl As Boolean
    SwDic As Dictionary
    Remaining As SwLins
End Type
Private Function EvlSwLins(A As SwLins, SwDic As Dictionary, Pm As Dictionary) As EvlSwLinsRslt
Dim J%, HasEvl As Boolean, Remaining As SwLins, M As SwLin
For J = 0 To A.N - 1
    M = A.Ay(J)
    With EvlSwLin(M, SwDic, Pm)
        If .Som Then
            SwDic.Add M.Nm, .Bool
            HasEvl = True
        Else
            PushSwLin Remaining, M
        End If
    End With
Next
EvlSwLins = EvlSwLinsRslt(HasEvl, Remaining, SwDic)
End Function
Private Function EvlSwLinsRslt(HasEvl As Boolean, Remaining As SwLins, SwDic As Dictionary) As EvlSwLinsRslt
With EvlSwLinsRslt
.HasEvl = HasEvl
.Remaining = Remaining
Set .SwDic = SwDic
End With
End Function
Private Function EvlSwLin(A As SwLin, SwDic As Dictionary, Pm As Dictionary) As BoolOpt
Static J&: J = J + 1
Dim O As BoolOpt
With A
Select Case True
Case .Op = EiOpAnd, .Op = EiOpOr: O = EvlSwLinAndOr(A.Op, .TermAy, SwDic, Pm)
Case .Op = EiOpEq, .Op = EiOpNe: O = EvlSwLinEqNe(A.Op, .T1, .T2, SwDic, Pm)
End Select
End With
EvlSwLin = O
If A.Nm = "?LvlY" And J = 3 Then Stop

End Function
Private Function EvlSwLinEqNe(Op As EmSwLinOp, T1$, T2$, SwDic As Dictionary, Pm As Dictionary) As BoolOpt
'Return True and set ORslt if evaluated
Dim S1$
    With EvlSwLinT1(T1, Pm)
        If Not .Som Then Exit Function
        S1 = .Str
    End With
Dim S2$
    With EvlSwLinT2(T2, Pm)
        If Not .Som Then Exit Function
        S2 = .Str
End With
Select Case True
Case Op = EiOpEq: EvlSwLinEqNe = SomBool(S1 = S2)
Case Op = EiOpNe: EvlSwLinEqNe = SomBool(S1 <> S2)
Case Else: ThwImpossible CSub
End Select
End Function
Private Function EvlSwLinAndOr(Op As EmSwLinOp, TermAy$(), SwDic As Dictionary, Pm As Dictionary) As BoolOpt
Dim J%, O() As Boolean
For J = 0 To UB(TermAy)
    With EvlSwTerm(TermAy(J), SwDic, Pm)
        If Not .Som Then Exit Function
        PushI O, .Bool
    End With
Next
Select Case True
Case Op = EiOpAnd: EvlSwLinAndOr = SomBool(IsAllTruezB(O))
Case Op = EiOpOr: EvlSwLinAndOr = SomBool(IsSomTruezB(O))
Case Else: ThwImpossible CSub
End Select
End Function
Private Function EvlSwTerm(SwTerm$, SwDic As Dictionary, Pm As Dictionary) As BoolOpt
Select Case True
Case SwDic.Exists(SwTerm): EvlSwTerm = SomBool(SwDic(SwTerm))
Case Pm.Exists(SwTerm):    EvlSwTerm = SomBool(Pm(SwTerm))
End Select
End Function

Private Function SwLinStr$(A As SwLin)
With A
Dim X$
Select Case True
Case .Op = EiOpAnd, .Op = EiOpOr: X = JnSpc(.TermAy)
Case .Op = EiOpEq, .Op = EiOpNe: X = .T1 & " " & .T2
End Select
SwLinStr = JnSpcAp(.Nm, SwLinOpStr(.Op), X)
End With
End Function
Private Function SwLinOpStr$(A As EmSwLinOp)
Dim O$
Select Case True
Case A = EiOpAnd: O = "And"
Case A = EiOpOr: O = "Or"
Case A = EiOpEq: O = "Eq"
Case A = EiOpNe: O = "Ne"
Case Else: ThwImpossible CSub
End Select
SwLinOpStr = O
End Function
Private Function SwLin(Nm$, Op As EmSwLinOp, T1$, T2$, TermAy$()) As SwLin
With SwLin
    .Nm = Nm
    .Op = Op
    Select Case True
    Case Op = EiOpNe, Op = EiOpEq: .T1 = T1: .T2 = T2
    Case Op = EiOpAnd, Op = EiOpOr: .TermAy = TermAy
    Case Else: Thw CSub, "Invalid Op", "Op", Op
    End Select
End With
End Function
Private Function SwLinOp(OpStr$) As EmSwLinOp
Select Case True
Case OpStr = "And": SwLinOp = EiOpAnd
Case OpStr = "Or": SwLinOp = EiOpOr
Case OpStr = "Eq": SwLinOp = EiOpEq
Case OpStr = "Ne": SwLinOp = EiOpNe
Case Else: Thw CSub, "Invalid OpStr", "OpStr VdtStr", OpStr, "And Or Eq Ne"
End Select
End Function
Private Function SwLinzLin(Lin) As SwLin
Dim Ay$(): Ay = TermAy(Lin)
Dim Nm$: Nm = Ay(0)
Dim OpStr$: OpStr = Ay(1)
Dim Op As EmSwLinOp: Op = SwLinOp(OpStr)
Dim T1$, T2$
Select Case True
Case Op = EiOpNe, Op = EiOpEq:
    If Si(Ay) <> 4 Then Thw CSub, "Lin should have 4 terms for Eq | Ne", "Lin", Lin
    T1 = Ay(2): T2 = Ay(3):
Case Op = EiOpAnd, Op = EiOpOr
    If Si(Ay) < 3 Then Thw CSub, "Lin should have at 3 terms And | Or", "Lin", Lin
    Ay = AyeFstNEle(Ay, 2)
End Select
SwLinzLin = SwLin(Nm, Op, T1, T2, Ay)
End Function
Private Sub PushSwLin(O As SwLins, M As SwLin)
ReDim Preserve O.Ay(O.N)
O.Ay(O.N) = M
O.N = O.N + 1
End Sub

Private Sub ZZ()
ZZ_SwLins
ZZ_SwzLyPm
End Sub
Private Sub ZZ_SwLins()
Dim SwLy$()
GoSub ZZ
Exit Sub
ZZ:
    BrwSwLins SwLins(Y_SwLy)
    Return
End Sub
Private Function SwLins(SwLy$()) As SwLins
Dim L
For Each L In Itr(SwLy)
    PushSwLin SwLins, SwLinzLin(L)
Next
End Function
Private Function Sw(StmtSw As Dictionary, FldSw As Dictionary) As Sw
Set Sw.StmtSw = StmtSw
Set Sw.FldSw = FldSw
End Function
Private Function SwzLyPm(SwLy$(), Pm As Dictionary) As Sw
Dim FldSwLy$():              FldSwLy = AyePfx(SwLy, "?:")
Dim StmtSwLy$():            StmtSwLy = AywPfx(SwLy, "?:")
Dim StmtSw As Dictionary: Set StmtSw = EvlSwLy(StmtSwLy, Pm)
Dim FldSw As Dictionary:  Set FldSw = EvlSwLy(FldSwLy, Pm)
SwzLyPm = Sw(StmtSw, FldSw)
End Function

Private Function EvlSwLy(SwLy$(), Pm As Dictionary) As Dictionary
Dim A As SwLins:          A = SwLins(SwLy)
Dim SwDic As New Dictionary
Dim R As EvlSwLinsRslt, J%
If Not Pm.Exists(">>SumLvl") Then Stop
Again:
    R = EvlSwLins(A, SwDic, Pm)
    J = J + 1: If J > 1000 Then ThwLoopingTooMuch CSub
    If R.Remaining.N = 0 Then
        Set EvlSwLy = R.SwDic
        Exit Function
    End If
    If R.HasEvl Then
        A = R.Remaining
        Set SwDic = R.SwDic
        GoTo Again
    End If
Thw CSub, "Cannot eval all StmtSwLins", "looping-cnt SwLy Pm Semi-Finished-SwDic Remaining-SwLins", J, SwLy, Pm, SwDic, FmtSwLins(R.Remaining)
End Function


Private Function EvlSwLinT1(T1$, Pm As Dictionary) As StrOpt
If Pm.Exists(T1) Then EvlSwLinT1 = SomStr(Pm(T1))
End Function

Private Function EvlSwLinT2(T2$, Pm As Dictionary) As StrOpt
If T2 = "*Blank" Then EvlSwLinT2 = SomStr(""): Exit Function
Dim M As StrOpt: M = EvlSwLinT1(T2, Pm)
If M.Som Then EvlSwLinT2 = M: Exit Function
EvlSwLinT2 = SomStr(T2)
End Function

Private Function EvlSwLinTerm(SwTerm$, Sw As Dictionary, Pm As Dictionary) As BoolOpt
Select Case True
Case Pm.Exists(SwTerm): EvlSwLinTerm = SomBool(Pm(SwTerm))
Case Sw.Exists(SwTerm): EvlSwLinTerm = SomBool(Sw(SwTerm))
End Select
End Function

Private Sub BrwSwLins(A As SwLins)
Brw FmtSwLins(A)
End Sub
Private Function FmtSwLins(A As SwLins) As String()
Dim J&
PushI FmtSwLins, "SwLins-Cnt: " & A.N
For J = 0 To A.N - 1
    PushI FmtSwLins, SwLinStr(A.Ay(J))
Next
End Function

Private Sub ZZ_SwzLyPm()
Dim SwLy$(), Pm As Dictionary
GoSub ZZ
Exit Sub
ZZ:
    BrwSw SwzLyPm(Y_SwLy, Y_Pm)
    Return
End Sub

Private Sub BrwSw(A As Sw)
B FmtSw(A)
End Sub
Private Function FmtSw(A As Sw) As String()
PushI FmtSw, "== StmtSw =================================="
PushIAy FmtSw, FmtDic(A.StmtSw)
PushI FmtSw, "== FldSw =================================="
PushIAy FmtSw, FmtDic(A.FldSw)
End Function

Private Function Y_SwLy() As String()
Erase XX
X "?LvlY    EQ >>SumLvl Y"
X "?LvlM    EQ >>SumLvl M"
X "?LvlW    EQ >>SumLvl W"
X "?LvlD    EQ >>SumLvl D"
X "?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY"
X "?M       OR ?LvlD ?LvlW ?LvlM"
X "?W       OR ?LvlD ?LvlW"
X "?D       OR ?LvlD"
X "?Dte     OR ?LvlD"
X "?Mbr     OR >?BrkMbr"
X "?MbrCnt  OR >?BrkMbr"
X "?Div     OR >?BrkDiv"
X "?Sto     OR >?BrkSto"
X "?Crd     OR >?BrkCrd"
X "?#SEL#Div NE >LisDiv *blank"
X "?#SEL#Sto NE >LisSto *blank"
X "?#SEL#Crd NE >LisCrd *blank"
Y_SwLy = XX
Erase XX
End Function
Private Sub ZZZ()
QTp_SqTp_Sw.SwLinOpStr
End Sub


Private Sub ZZ_SqyzTp()
Dim SqTp$
GoSub ZZ
Exit Sub
ZZ:
    B SqyzTp(Y_SqTp)
    Return
End Sub
Function SqyzTp(SqTp$) As String()
'ThwIf_Er ErzSqTp(SqTp), CSub
Dim B As Blks:             B = BlkszSqTp(SqTp)
Dim Pm As Dictionary: Set Pm = PmzLy(LyzBlksTy(B, "PM"))
Dim Sw As Sw:             Sw = SwzLyPm(LyzBlksTy(B, "SW"), Pm)
SqyzTp = SqyzLyPmSw(LyAyzBlksTy(B, "SQ"), Pm, Sw)
End Function
Function PmzLy(PmLy$()) As Dictionary
Dim O As Dictionary: Set O = Dic(PmLy)
Dim B As Boolean, K, V
For Each K In O.Keys
    Select Case True
    Case HasPfx(K, ">?"):
        V = O(K)
        Select Case V
        Case "0": B = False
        Case "1": B = True
        Case Else: Thw CSub, "If K='>?xxx', V should be 0 or 1", "K V PmLy", K, V, PmLy
        End Select
        O(K) = B
    Case HasPfx(K, ">")
    Case Else: Thw CSub, "Pm line should beg with (>? | >)", "K V PmLy", K, O(K), PmLy
    End Select
Next
Set PmzLy = O
End Function
Private Sub ZZ_PmzLy()
B Y_PmLy
End Sub

Private Sub ZZ_Pm()
BrwDic Y_Pm
End Sub
Sub Z3()
ZZ_SqyzTp
End Sub

Private Function Y_Blks() As Blks
Y_Blks = BlkszSqTp(Y_SqTp)
End Function
Private Function Y_PmLy() As String()
Y_PmLy = LyzBlksTy(Y_Blks, "PM")
End Function

Private Function SqyzLyPmSw(SqLyAy(), Pm As Dictionary, Sw As Sw) As String()
Dim SqLy
For Each SqLy In Itr(SqLyAy)
    PushI SqyzLyPmSw, SqlzLyPmSw(CvSy(SqLy), Pm, Sw)
Next
End Function

Private Function BlkszSqTp(SqTp$) As Blks
Dim O As Blks, J&
O = Blks(SqTp)
For J = 0 To O.N - 1
    O.Ay(J).BlkTy = BlkTy(LinAyzLnxs(O.Ay(J).Lnxs))
Next
BlkszSqTp = O
End Function

Private Function IsPm(Ly$()) As Boolean
IsPm = HasPfxOfAllEle(Ly, ">")
End Function

Private Function IsRm(Ly$()) As Boolean
IsRm = Si(Ly) = 0
End Function
Private Sub ZZ_IsSq()
Dim Ly$()
GoSub ZZ
Exit Sub
ZZ:
    Debug.Assert IsSqLy(Sy("Drp sdf"))
    Return
End Sub
Private Function IsSq(Ly$()) As Boolean
If Si(Ly) = 0 Then Exit Function
Dim L$: L = Ly(0)
Dim Sy$(): Sy = SyzSS("?SEL SEL ?SELDIS SELDIS UPD DRP")
If HitPfxAySpc(L, Sy, vbTextCompare) Then IsSq = True: Exit Function
End Function

Private Function IsSw(Ly$()) As Boolean
IsSw = HasPfxOfAllEle(Ly, "?")
End Function
Private Function Y_Pm() As Dictionary
Set Y_Pm = PmzLy(Y_PmLy)
End Function

Private Property Get Y_SwLnxs() As Lnxs
End Property
Private Property Get Y_Sw() As Dictionary
End Property
Private Property Get Y_FldSw() As Dictionary

End Property
Private Property Get Y_StmtSw() As Dictionary

End Property
Private Function Y_SqTp$()
Y_SqTp = SampSqTp
End Function
Property Get SampSqTp$()
Erase XX
X "-- Rmk: -- is remark"
X "-- >XX: is PmLin"
X "-- >?XX: is switchPrm, it value must be 0 or 1"
X "-- ?XX: is switch line"
X "-- SwitchLin: is ?XXX [OR|AND|EQ|NE] [SwPrm_OR_AND|SwPrm_EQ_NE]"
X "-- SwPrm_OR_AND: SwTerm .."
X "-- SwPrm_EQ_NE:  SwEQ_NE_T1 SwEQ_NE_T2"
X "-- SwEQ_NE_T1:"
X "-- SwEQ_NE_T2:"
X "-- SwTerm:     ?XX|>?XX     -- if >?XX, its value only 1 or 0 is allowed"
X "-- Only one gp of >XX:"
X "-- Only one gp of ?XX:"
X "-- All other gp is sql-statement or sql-statements"
X "-- sql-statments: Drp xxx xxx"
X "-- sql-statment: [sel|selDis|upd|into|fm|whBetStr|whBetNbr|whInStrLis|whInNbrLis|andInNbrLis|andInStrLis|gp|jn|left|expr]"
X "-- optional: Whxxx and Andxxx can have ?-pfx becomes: ?Whxxx and ?Andxxx.  The line will become empty"
X "=============================================="
X "Drp Tx TxMbr MbrDta Div Sto Crd Cnt Oup MbrWs"
X "============================================="
X "-- >?XXX means input switch, value must be 0 or 1"
X "-- >XXX  means input txt and optional, allow, blank"
X "-- >>XXX means input compulasary, that means not allow blank"
X ">?BrkMbr 0"
X ">?BrkSto 0"
X ">?BrkCrd 0"
X ">?BrkDiv 0"
X ">>SumLvl  Y"
X ">?MbrEmail 1"
X ">?MbrNm    1"
X ">?MbrPhone 1"
X ">?MbrAdr   1"
X ">>DteFm 20170101"
X ">>DteTo 20170131"
X ">LisDiv 1 2"
X ">LisSto"
X ">LisCrd"
X ">CrdExpr ..."
X ">CrdExpr ..."
X ">CrdExpr ..."
X "============================================"
X "-- EQ & NE t1 only TxtPm is allowed"
X "--         t2 allow TxtPm, *BLANK, and other text"
X "?LvlY    EQ >>SumLvl Y"
X "?LvlM    EQ >>SumLvl M"
X "?LvlW    EQ >>SumLvl W"
X "?LvlD    EQ >>SumLvl D"
X "?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY"
X "?M       OR ?LvlD ?LvlW ?LvlM"
X "?W       OR ?LvlD ?LvlW"
X "?D       OR ?LvlD"
X "?WD      OR ?LvlD"
X "?Dte     OR ?LvlD"
X "?Mbr     OR >?BrkMbr"
X "?MbrCnt  OR >?BrkMbr"
X "?Div     OR >?BrkDiv"
X "?Sto     OR >?BrkSto"
X "?Crd     OR >?BrkCrd"
X "?:#Div NE >LisDiv *blank"
X "?:#Sto NE >LisSto *blank"
X "?:#Crd NE >LisCrd *blank"
X "============================================= #Tx"
X "sel  ?Crd ?Mbr ?Div ?Sto ?Y ?M ?W ?WD ?D ?Dte Amt Qty Cnt"
X "into #Tx"
X "fm   SalesHistory"
X "wh   bet str    >>DteFm >>DteTo"
X "?and in  strlis Div >LisDiv"
X "?and in  strlis Sto >LisSto"
X "?and in  nbrlis Crd >LisCrd"
X "?gp  ?Crd ?Mbr ?Div ?Sto ?Crd ?Y ?M ?W ?WD ?D ?Dte"
X "$Crd >CrdExpr"
X "$Mbr JCMCode"
X "$Sto"
X "$Y"
X "$M"
X "$W"
X "$WD"
X "$D"
X "$Dte"
X "$Amt Sum(SHAmount)"
X "$Qty Sum(SHQty)"
X "$Cnt Count(SHInvoice+SHSDate+SHRef)"
X "============================================= #TxMbr"
X "selDis  Mbr"
X "fm      #Tx"
X "into    #TxMbr"
X "============================================= #MbrDta"
X "sel   Mbr Age Sex Sts Dist Area"
X "fm    #TxMbr x"
X "jn    JCMMember a on x.Mbr = a.JCMMCode"
X "into  #MbrDta"
X "$Mbr  x.Mbr"
X "$Age  DATEDIFF(YEAR,CONVERT(DATETIME ,x.JCMDOB,112),GETDATE())"
X "$Sex  a.JCMSex"
X "$Sts  a.JCMStatus"
X "$Dist a.JCMDist"
X "$Area a.JCMArea"
X "==-=========================================== #Div"
X "?sel Div DivNm DivSeq DivSts"
X "fm   Division"
X "into #Div"
X "?wh in strLis Div >LisDiv"
X "$Div Dept + Division"
X "$DivNm LongDies"
X "$DivSeq Seq"
X "$DivSts Status"
X "============================================ #Sto"
X "?sel Sto StoNm StoCNm"
X "fm   Location"
X "into #Sto"
X "?wh in strLis Loc >LisLoc"
X "$Sto"
X "$StoNm"
X "$StoCNm"
X "============================================= #Crd"
X "?sel        Crd CrdNm"
X "fm          Location"
X "into        #Crd"
X "?wh in nbrLis Crd >LisCrd"
X "$Crd"
X "$CrdNm"
X "============================================= #Oup"
X "sel  ?Crd ?CrdNm ?Mbr ?Age ?Sex ?Sts ?Dist ?Area ?Div ?DivNm ?Sto ?StoNm ?StoCNm ?Y ?M ?W ?WD ?D ?Dte Amt Qty Cnt"
X "into #Oup"
X "fm   #Tx x"
X "left #Crd a on x.Crd = a.Crd"
X "left #Div b on x.Div = b.Div"
X "left #Sto c on x.Sto = c.Sto"
X "left #MbrDta d on x.Mbr = d.Mbr"
X "wh   JCMCode in (Select Mbr From #TxMbr)"
X "============================================ #Cnt"
X "sel ?MbrCnt RecCnt TxCnt Qty Amt"
X "into #Cnt"
X "fm  #Tx"
X "$MbrCnt?Count(Distinct Mbr)"
X "$RecCnt Count(*)"
X "$TxCnt  Sum(TxCnt)"
X "$Qty    Sum(Qty)"
X "$Amt    Sum(Amt)"
X "============================================"
X "--"
X "============================================"
X "df eror fs--"
X "============================================"
X "-- EQ & NE t1 only TxtPm is allowed"
X "--         t2 allow TxtPm, *BLANK, and other text"
X "?LvlY    EQ >>SumLvl Y"
X "?LvlM    EQ >>SumLvl M"
X "?LvlW    EQ >>SumLvl W"
X "?LvlD    EQ >>SumLvl D"
X "?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY`"
SampSqTp = JnCrLf(XX)
Erase XX
End Property

Private Function BlkTy$(Ly$())
Dim O$
Select Case True
Case IsPm(Ly): O = "PM"
Case IsSw(Ly): O = "SW"
Case IsRm(Ly): O = "RM"
Case IsSq(Ly): O = "SQ"
'Case Else: Brw Ly: Stop: O = "ER"
Case Else:        O = "ER"
End Select
BlkTy = O
End Function

Function SqlSel_X_Into_T_Wh_Gp_Ord$(X$, Into$, T$, Gp$, Ord$)

End Function

Private Function SQ_ExprLis$(Fny$(), EDic As Dictionary, FzDiAlias As Dictionary)
SqpSelX_Fny_ExtNy_ODis
End Function
Private Function SqlSel$(Sel As KLys, EDic As Dictionary, FldSw As Dictionary)
'BrwKLys Sel
Dim X As KLys: X = Sel
Dim LFm$, LInto$, LSel$, LOrd$, LWh$, LGp$, LAndOr$(), LAlias$()
'    LSel = ShfKLyMLin(X, "Sel")
'    LInto = ShfKLyMLin(X, "Into")
'    LFm = ShfKLyMLin(X, "Fm")
'    LJn = ShfKLyMLyzKK(X, "Jn LJn")
'    LWh = ShfKLyOLin(X, "Wh")
'    LAndOr = ShfKLyMLyzKK(X, "And Or")
'    LGp = ShfKLyOLin(X, "Gp")
'    LOrd = ShfKLyOLin(X, "Ord")
Dim ADic As Dictionary: Set ADic = DiczVkkLy(LAlias)
Dim Ffny$(), FGp$()
    Ffny = SQ_SelFny(LSel, FldSw)
    FGp = SQ_SelFld(LGp, FldSw)
Dim OX$, OInto$, OT$, OWh$, OGp$, OOrd$
'    Dim Fny$()

'    '
    OInto = RmvT1(LInto)
    OGp = SQ_ExprLis(FGp, EDic, ADic)
    OOrd = SQ_ExprLis(FOrd, EDic, ADic)
'    OWh = SQ_Wh()
'    OT = RmvT1(LFm)
SqlSel = SqlSel_X_Into_T_Wh_Gp_Ord(OX, OInto, OT, OGp, OOrd)
End Function

Private Function SQ_SelFld(FF$, FldSw As Dictionary) As String()
Dim Fny$(): Fny = SyzSS(FF)
Dim F1$, F
For Each F In Fny
    F1 = FstChr(F)
    Select Case True
    Case F1 = "?"
        If Not FldSw.Exists(F) Then Thw CSub, "An option fld not found in FldSw", "Opt-Fld FF FldSw", F, FF, FldSw
        If FldSw(F) Then
            PushI XFny, RmvFstChr(F)
        End If
    Case F1 = "$"
        PushI XFny, RmvFstChr(F)
    Case Else
        PushI XFny, F
    End Select
Next
Stop
End Function

Private Function SqlUpd$(Upd As KLys, EDic As Dictionary, FldSw As Dictionary)
End Function

Private Function SqlDrp$(SqLy$())
End Function

Private Function IsSkip(FstSqLin$, SqLy$(), Ty As EmStmtTy, StmtSw As Dictionary) As Boolean
If FstChr(FstSqLin) <> "?" Then Exit Function
Dim Key$: Key = StmtSwKey(SqLy, Ty)
If Not StmtSw.Exists(Key) Then Thw CSub, "StmtSw does not contain the StmtSwKey", "SqLy StmtSwKey StmtSw", SqLy, Key, StmtSw
IsSkip = Not StmtSw(Key)
End Function

Private Function SqlzLyPmSw$(SqLy$(), Pm As Dictionary, Sw As Sw)
Dim FstSqLin$:    FstSqLin = SqLy(0)
Dim Ty As EmStmtTy:     Ty = StmtTy(FstSqLin)
Dim Skip As Boolean:  Skip = IsSkip(FstSqLin, SqLy, Ty, Sw.StmtSw)
                             If Skip Then Exit Function
Dim S$():                S = SQ_RmvExprLin(SqLy)
Dim E As Dictionary: Set E = SQ_ExprDic(SqLy)
Dim O$
    Select Case True
    Case Ty = EiDrpStmt: O = SqlDrp(S)
    Case Ty = EiUpdStmt: O = SqlUpd(KLys(S), E, Sw.FldSw)
    Case Ty = EiSelStmt: O = SqlSel(KLys(S), E, Sw.FldSw)
    Case Else: ThwImpossible CSub
    End Select
SqlzLyPmSw = O
End Function

Private Function SQ_RmvExprLin(SqLy$()) As String()
SQ_RmvExprLin = AyePfx(SqLy, "$")
End Function


Private Function SQ_ExprDic(SqLy$()) As Dictionary
Set SQ_ExprDic = Dic(CvSy(AywPfx(SqLy, "$")))
End Function

Private Function StmtTy(FstSqLin$) As EmStmtTy
Dim L$: L = RmvPfx(T1(FstSqLin), "?")
Select Case L
Case "SEL", "SELDIS": StmtTy = EiSelStmt
Case "UPD": StmtTy = EiUpdStmt
Case "DRP": StmtTy = EiDrpStmt
Case Else: Stop
End Select
End Function

Private Function StmtSwKey$(SqLy$(), Ty As EmStmtTy)
Dim O$
Select Case Ty
Case EiSelStmt: O = StmtSwKeyzSel(SqLy)
Case EiUpdStmt: O = StmtSwKeyzUpd(SqLy)
Case Else: 'Only Sel/Upd can have StmtSwKey
End Select
StmtSwKey = "?:" & O
End Function

Private Function StmtSwKeyzSel$(SelSqLy$())
StmtSwKeyzSel = FstElewRmvT1(SelSqLy, "into")
End Function

Private Function StmtSwKeyzUpd$(UpdSqLy$())
Dim Lin1$
    Lin1 = UpdSqLy(0)
If RmvPfx(ShfT1(Lin1), "?") <> "upd" Then Stop
StmtSwKeyzUpd = Lin1
End Function


Private Function IsXXX(A$(), XXX$) As Boolean
IsXXX = UCase(T1(A(UB(A)))) = XXX
End Function


Private Property Get Y_ExprDic() As Dictionary
Dim O$()
PushI O, "A XX"
PushI O, "B BB"
PushI O, "C DD"
PushI O, "E FF"
'Set Y_ExprDic = LyDic(O)
End Property

Private Property Get Y_SqLnxs() As Lnxs
Dim O$()
PushI O, "sel ?MbrCnt RecCnt TxCnt Qty Amt"
PushI O, "into #Cnt"
PushI O, "fm   #Tx"
PushI O, "wh   RecCnt bet @XX @XX"
PushI O, "and  RecCnt bet @XX @XX"

PushI O, "$"
PushI O, "?MbrCnt ?Count(Distinct Mbr)"
PushI O, "RecCnt  Count(*)"
PushI O, "TxCnt   Sum(TxCnt)"
PushI O, "Qty     Sum(Qty)"
PushI O, "Amt     Sum(Amt)"
Y_SqLnxs = Lnxs(O)
End Property

Private Function SQ_And(A$(), E As Dictionary)
'and f bet xx xx
'and f in xx
Dim F$, I, L$, Ix%, M As Lnx
For Each I In Itr(A)
    'Set M = I
    LnxAsg M, L, Ix
    If ShfT1(L) <> "and" Then Stop
    F = ShfT1(L)
    Select Case ShfT1(L)
    Case "bet":
    Case "in"
    Case Else: Stop
    End Select
Next
End Function

Private Function SQ_Gp$(GG$, FldSw As Dictionary, E As Dictionary)
If GG = "" Then Exit Function
Dim ExprAy$(), Ay$()
Stop
'    ExprAy = DicSelIntoSy(EDic, Ay)
'XGp = SqpGp(ExprAy)
End Function

Private Function SQ_JnOrLeftJn(A$(), E As Dictionary) As String()

End Function

Private Function PopJnOrLeftJn(A$()) As String()
PopJnOrLeftJn = PopMulXorYOpt(A, U_Jn, U_LeftJn)
End Function

Private Function PopXXXOpt$(A$(), XXX$)
'Return No-Q-T1-of-LasEle-of-A$() if No-Q-T1-of it = XXX else return ''
If Si(A) = 0 Then Exit Function
PopXXXOpt = PopXXX(A, XXX)
End Function

Private Function SQ_Sel$(A$, E As Dictionary)
Dim Fny$()
    Dim T1$, L$
    L = A
    T1 = RmvPfx(ShfT1(L), "?")
    'Fny = XSelFny(SyzSS(L), FldSw)
Select Case T1
'Case U_Sel:    XSel = X.Sel_Fny_EDic(Fny, E)
'Case U_SelDis: XSel = X.Sel_Fny_EDic(Fny, E, IsDis:=True)
Case Else: Stop
End Select
End Function
Private Function SQ_SelFny(Fny$(), FldSw As Dictionary) As String()
Dim F
For Each F In Fny
    If FstChr(F) = "?" Then
        If Not FldSw.Exists(F) Then Stop
        If FldSw(F) Then PushI XSelFny, F
    Else
        PushI XSelFny, F
    End If
Next
End Function

Private Function SQ_Set(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function SQ_Upd(A As Lnx, E As Dictionary, OEr$())

End Function
Private Function SQ_Wh$() ' (L$, E As Dictionary)
'L is following
'  ?Fld in @ValLis  -
'  ?Fld bet @V1 @V2
Dim F$, Vy$(), V1, V2, IsBet As Boolean
If IsBet Then
'    If Not FndValPair(F, E, V1, V2) Then Exit Function
    'XWh = SqpWhBet(F, V1, V2)
    Exit Function
End If
'If Not FndVy(F, E, Vy, Q) Then Exit Function
'XWh = SqpWhFldInVy_Str(F, Vy)
End Function

Private Function SQ_WhBetNbr$(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function SQ_WhExpr(A As Lnx, E As Dictionary, OEr$())

End Function

Private Function SQ_WhInNbrLis$(A As Lnx, E As Dictionary, OEr$())

End Function

Private Sub Z_SqlRsltSel()
Dim E As Dictionary, Ly$(), Act As SqlRslt
'---
Erase Ly
    Push Ly, "?XX Fld-XX"
    Push Ly, "BB Fld-BB-LINE-1"
'    Push Ly, "BB Fld-BB-LINE-2"
'    Set E = LyDic(Ly)           '<== Set ExprDic
Erase Ly
    'Set FldSw = New Dictionary
    'FldSw.Add "?XX", False       '<=== Set FldSw
Erase Ly
    Erase Ly
    PushI Ly, "sel ?XX BB CC"
    PushI Ly, "into #AA"
    PushI Ly, "fm   #AA"
    PushI Ly, "jn   #AA"
    PushI Ly, "jn   #AA"
    PushI Ly, "wh   A bet $a $b"
    PushI Ly, "and  B in $c"
    PushI Ly, "gp   D C"        '<== LySq
GoSub Tst
Exit Sub
Tst:
    Act = SqlRsltSel(Ly, E)
    C
    Return
End Sub

Private Sub Z_ExprDic()
Dim Ly$()
Dim D As New Dictionary
'-----

Erase Ly
PushI Ly, "aaa bbb"
PushI Ly, "111 222"
PushI Ly, "$"
PushI Ly, "A B0"
PushI Ly, "A B1"
PushI Ly, "A B2"
PushI Ly, "B B0"
D.RemoveAll
    D.Add "A", JnCrLf(SyzSS("B0 B1 B2"))
    D.Add "B", "B0"
    Set Ept = D
GoSub Tst
Exit Sub
Tst:
    Set Act = SQ_ExprDic(Ly)
    Ass IsEqDic(CvDic(Act), CvDic(Ept))
    
    Return
End Sub
Private Sub ZZ_StmtSwKey()
Dim Ly$(), Ty As EmStmtTy
GoSub T0
GoSub T1
Exit Sub
'---
T0:
    Erase Ly
    PushI Ly, "sel sdflk"
    PushI Ly, "fm AA BB"
    PushI Ly, "into XX"
    Ept = "XX"
    Ty = EiSelStmt
    GoTo Tst
T1:
    Erase Ly
    PushI Ly, "?upd XX BB"
    PushI Ly, "fm dsklf dsfl"
    Ept = "XX BB"
    Ty = EiUpdStmt
    GoTo Tst
Tst:
    Act = StmtSwKey(Ly, Ty)
    C
    Return
End Sub

