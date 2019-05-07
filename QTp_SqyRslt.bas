Attribute VB_Name = "QTp_SqyRslt"
Option Explicit
Private Const CMod$ = "MTp_SqyRslt."
Private Const Asm$ = "QTp"
Type SqyRslt: Er() As String: Sqy() As String: End Type
Enum EmSqTpBlkTy
    EiErBlk
    EiPmBlk
    EiSwBlk
    EiSqBlk
    EiRmBlk
End Enum
Const SqBlkTyNN$ = "ER PM SW SQ RM"
Public SampSqt As New SampSqt
Function SqyRsltzEr(Sqy$(), Er$()) As SqyRslt
SqyRsltzEr.Er = Er
SqyRsltzEr.Sqy = Sqy
End Function

Function SqyRslt(SqTp$) As SqyRslt
Dim B() As Blk:            B = BlkAy(SqTp)
Dim PmR As PmRslt:       PmR = PmRsltzLnxAy(LnxAyzBlk(B, "PM"))
Dim Pm As Dictionary: Set Pm = PmR.Pm
Dim SwR As SwRslt:       SwR = SwRsltzLnxAy(LnxAyzBlk(B, "SW"), Pm)
Dim SqR As SqyRslt:      SqR = SqyRsltzGpAy(GpAyzBlkTy(B, "SQ"), Pm, SwR.StmtSw, SwR.FldSw)
Dim Er$():                Er = AddAyAp(ErzBlkAy(B), PmR.Er, SwR.Er, SqR.Er)
                     SqyRslt = SqyRsltzEr(SqR.Sqy, Er)
End Function

Private Function LnxAyzBlk(A() As Blk, BlkTy$) As Lnx()
Dim J%
For J = 0 To UB(A)
    If A(J).BlkTy = BlkTy Then LnxAyzBlk = A(J).Gp.LnxAy: Exit Function
Next
End Function

Private Property Get PmLy() As String()
PushI PmLy, "@?BrkMbr 0"
PushI PmLy, "@?BrkSto 0"
PushI PmLy, "@?BrkCrd 0"
PushI PmLy, "@?BrkDiv 0"
'-- @XXX means txt and optional, allow, blank
PushI PmLy, "@SumLvl  Y"
PushI PmLy, "@?MbrEmail 1"
PushI PmLy, "@?MbrNm    1"
PushI PmLy, "@?MbrPhone 1"
PushI PmLy, "@?MbrAdr   1"
'-- @@ mean compulasary
PushI PmLy, "@DteFm 20170101"
'@@DteTo 20170131
PushI PmLy, "@LisDiv 1 2"
PushI PmLy, "@LisSto"
PushI PmLy, "@LisCrd"
PushI PmLy, "@CrdExpr ..."
PushI PmLy, "@CrdExpr ..."
PushI PmLy, "@CrdExpr ..."
End Property
Property Get SampPm() As Dictionary
Set Pm = Dic(PmLy, vbCrLf)
End Property

Property Get SampSwLnxAy() As Lnx()
Erase XX
X "?#LvlY    EQ @SumLvl Y"
X "?#LvlY    EQ @SumLvl Y"
X "?#LvlM    EQ @SumLvl M"
X "?#LvlW    EQ @SumLvl W"
X "?#LvlD    EQ @SumLvl D"
X "?Y       OR ?#LvlD ?#LvlW ?#LvlM ?#LvlY"
X "?M       OR ?#LvlD ?#LvlW ?#LvlM"
X "?W       OR ?#LvlD ?#LvlW"
X "?D       OR ?#LvlD"
X "?Dte     OR ?#LvlD"
X "?Mbr     OR @?BrkMbr XX"
X "?MbrCnt  OR @?BrkMbr"
X "?Div     OR @?BrkDiv"
X "?Sto     OR @?BrkSto"
X "?Crd     OR @?BrkCrd"
X "?SEL#Div NE @LisDiv *blank"
X "?SEL#Sto NE @LisSto *blank"
X "?SEL#Crd NE @LisCrd *blank"
SwLnxAy = LnxAy(XX)
Erase XX
End Property
Property Get SampSw() As Dictionary

End Property
Property Get SampFldSw() As Dictionary

End Property
Property Get SampStmtSw() As Dictionary

End Property

Property Get SampSqTp$()
Erase XX
X "-- Rmk: -- is remark"
X "-- %XX: is prmDicLin"
X "-- %?XX: is switchPrm, it value must be 0 or 1"
X "-- ?XX: is switch line"
X "-- SwitchLin: is ?XXX [OR|AND|EQ|NE] [SwPrm_OR_AND|SwPrm_EQ_NE]"
X "-- SwPrm_OR_AND: SwTerm .."
X "-- SwPrm_EQ_NE:  SwEQ_NE_T1 SwEQ_NE_T2"
X "-- SwEQ_NE_T1:"
X "-- SwEQ_NE_T2:"
X "-- SwTerm:     ?XX|%?XX     -- if %?XX, its value only 1 or 0 is allowed"
X "-- Only one gp of %XX:"
X "-- Only one gp of ?XX:"
X "-- All other gp is sql-statement or sql-statements"
X "-- sql-statments: Drp xxx xxx"
X "-- sql-statment: [sel|selDis|upd|into|fm|whBetStr|whBetNbr|whInStrLis|whInNbrLis|andInNbrLis|andInStrLis|gp|jn|left|expr]"
X "-- optional: Whxxx and Andxxx can have ?-pfx becomes: ?Whxxx and ?Andxxx.  The line will become empty"
X "=============================================="
X "Drp Tx TxMbr MbrDta Div Sto Crd Cnt Oup MbrWs"
X "============================================="
X "-- @? means switch, value must be 0 or 1"
X "@?BrkMbr 0"
X "@?BrkMbr 0"
X "@?BrkMbr 0"
X "@?BrkSto 0"
X "@?BrkCrd 0"
X "@?BrkDiv 0"
X "-- %XXX means txt and optional, allow, blank"
X "@SumLvl  Y"
X "@?MbrEmail 1"
X "@?MbrNm    1"
X "@?MbrPhone 1"
X "@?MbrAdr   1"
X "-- %% mean compulasary"
X "@%DteFm 20170101"
X "@%DteTo 20170131"
X "@LisDiv 1 2"
X "@LisSto"
X "@LisCrd"
X "@CrdExpr ..."
X "@CrdExpr ..."
X "@CrdExpr ..."
X "============================================"
X "-- EQ & NE t1 only TxtPm is allowed"
X "--         t2 allow TxtPm, *BLANK, and other text"
X "?LvlY    EQ %SumLvl Y"
X "?LvlM    EQ %SumLvl M"
X "?LvlW    EQ %SumLvl W"
X "?LvlD    EQ %SumLvl D"
X "?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY"
X "?M       OR ?LvlD ?LvlW ?LvlM"
X "?W       OR ?LvlD ?LvlW"
X "?D       OR ?LvlD"
X "?Dte     OR ?LvlD"
X "?Mbr     OR %?BrkMbr"
X "?MbrCnt  OR %?BrkMbr"
X "?Div     OR %?BrkDiv"
X "?Sto     OR %?BrkSto"
X "?Crd     OR %?BrkCrd"
X "?#SEL#Div NE %LisDiv *blank"
X "?#SEL#Sto NE %LisSto *blank"
X "?#SEL#Crd NE %LisCrd *blank"
X "============================================= #Tx"
X "sel  ?Crd ?Mbr??Div ?Sto ?Y ?M ?W ?WD ?D ?Dte Amt Qty Cnt"
X "into #Tx"
X "fm   SalesHistory"
X "wh   bet str    %%DteFm %%DteTo"
X "?and in  strlis Div %LisDiv"
X "?and in  strlis Sto %LisSto"
X "?and in  nbrlis Crd %LisCrd"
X "?gp  ?Crd ?Mbr ?Div ?Sto ?Crd ?Y ?M ?W ?WD ?D ?Dte"
X "$Crd %CrdExpr"
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
X "?wh in strLis Div %LisDiv"
X "$Div Dept + Division"
X "$DivNm LongDies"
X "$DivSeq Seq"
X "$DivSts Status"
X "============================================ #Sto"
X "?sel Sto StoNm StoCNm"
X "fm   Location"
X "into #Sto"
X "?wh in strLis Loc %LisLoc"
X "$Sto"
X "$StoNm"
X "$StoCNm"
X "============================================= #Crd"
X "?sel        Crd CrdNm"
X "fm          Location"
X "into        #Crd"
X "?wh in nbrLis Crd %LisCrd"
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
X "?LvlY    EQ %SumLvl Y"
X "?LvlM    EQ %SumLvl M"
X "?LvlW    EQ %SumLvl W"
X "?LvlD    EQ %SumLvl D"
X "?Y       OR ?LvlD ?LvlW ?LvlM ?LvlY`"
SqTp = JnCrLf(XX)
Erase XX
End Property

