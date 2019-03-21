Attribute VB_Name = "MXls_Lof_ErzLof"
Option Explicit
Const CMod$ = ""
Public Const LofT1nn$ = _
                            "Lo Ali Bdr Tot Wdt Fmt Lvl Cor Fml Lbl Tit Bet" ' Fmt. i.tm s.pace s.eparated string
Public Const LofT1nnzSng$ = "                               Fml Lbl Tit Bet" ' Sng.sigle field per line
Public Const LofT1nnzMul$ = "Lo Ali Bdr Tot Wdt Fmt Lvl Cor                " ' Mul.tiple field per line
'GenErMsg-Src-Beg.
'Val_NotNum      Lno#{Lno&} is [{T1$}] line having Val({Val$}) which should be a number
'Val_NotBet      Lno#{Lno&} is [{T1$}] line having Val({Val$}) which between ({FmNo}) and (ToNm})
'Val_NotInLis    Lno#{Lno&} is [{T1$}] line having invalid Val({ErVal$}).  See valid-value-{VdtValNm$}
'Val_FmlFld      Lno#{Lno&} is [Fml] line having invalid Fml({Fml$}) due to invalid Fny({ErFny$()}).  Valid-Fny are [{VdtFny$()}]
'Val_FmlNotBegEq Lno#{Lno&} is [Fml] line having [{Fml$}] which is not started with [=]
'Fld_NotInFny    Lno#{Lno&} is [{T1$}] line having Fld({F}) which should one of the Fny value.  See [Fny-Value]
'Fld_Dup         Lno#{Lno&} is [{T1}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno&}
'Fldss_NotSel    Lno#{Lno&} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]
'Fldss_DupSel    Lno#{Lno&} is [{T1$}] line having
'LoNm            Lno#{Lno&} is [Lo-Nm] line having value({Val$}) which is not a good name.
'LoNm_Mis        [Lo-Nm] line is missing
'LoNm_Dup        Lno#{Lno&} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno&}
'Tot_DupSel      Lno#{Lno&} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno&} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored.
'Bet_3Fld        Lno#{Lno&} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]"
'Bet_EqFmTo      Lno#{Lno&} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal.
'Bet_FldSeq      Lno#{Lno&} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]"
'GenErMsg-Src-End.
Private A$(), A_Fny$(), A_T1ToLnxAyDic As Dictionary
Private Const M_Val_NotNum$ = "Lno#{Lno&} is [{T1$}] line having Val({Val$}) which should be a number"
Private Const M_Val_NotBet$ = "Lno#{Lno&} is [{T1$}] line having Val({Val$}) which between ({FmNo}) and (ToNm})"
Private Const M_Val_NotInLis$ = "Lno#{Lno&} is [{T1$}] line having invalid Val({ErVal$}).  See valid-value-{VdtValNm$}"
Private Const M_Val_FmlFld$ = "Lno#{Lno&} is [Fml] line having invalid Fml({Fml$}) due to invalid Fny({ErFny$()}).  Valid-Fny are [{VdtFny$()}]"
Private Const M_Val_FmlNotBegEq$ = "Lno#{Lno&} is [Fml] line having [{Fml$}] which is not started with [=]"
Private Const M_Fld_NotInFny$ = "Lno#{Lno&} is [{T1$}] line having Fld({F}) which should one of the Fny value.  See [Fny-Value]"
Private Const M_Fld_Dup$ = "Lno#{Lno&} is [{T1}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno&}"
Private Const M_Fldss_NotSel$ = "Lno#{Lno&} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]"
Private Const M_Fldss_DupSel$ = "Lno#{Lno&} is [{T1$}] line having"
Private Const M_LoNm$ = "Lno#{Lno&} is [Lo-Nm] line having value({Val$}) which is not a good name."
Private Const M_LoNm_Mis$ = "[Lo-Nm] line is missing"
Private Const M_LoNm_Dup$ = "Lno#{Lno&} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno&}"
Private Const M_Tot_DupSel$ = "Lno#{Lno&} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno&} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored."
Private Const M_Bet_3Fld$ = "Lno#{Lno&} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]"""
Private Const M_Bet_EqFmTo$ = "Lno#{Lno&} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal."
Private Const M_Bet_FldSeq$ = "Lno#{Lno&} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]"""
Private Function MsgVal_FmlNotBegWithEq$(Lno&, Fml$)

End Function
Property Get LofT1Ny() As String()
LofT1Ny = NyzNN(LofT1nn)
End Property
Function ErzLof(Lof$(), Fny$()) As String() 'Error-of-ListObj-Formatter:Er.z.Lo.f
Const CSub$ = CMod & "LofEr"
Init Lof, Fny
ErzLof = SyAddAp( _
    ErVal, ErFld, ErFldss, ErLoNm, _
    ErAli, ErBdr, ErTot, _
    ErWdt, ErFmt, ErLvl, ErCor, _
    ErFml, ErLbl, ErTit, ErBet)
End Function
Function FnywLikssAy(Fny$(), LikssAy$()) As String()
Dim F, LikAy$()
LikAy = TermAsetzTLinAy(LikssAy).Sy
For Each F In Itr(Fny)
    If HitLikAy(F, LikAy) Then PushI FnywLikssAy, F
Next
End Function
Private Sub Init(Lof$(), Fny$())
A = Lof
A_Fny = Fny
Set A_T1ToLnxAyDic = LnxAyDiczT1nn(A, LofT1nn)
End Sub
Private Property Get WAli_LeftRightCenter() As String()
'ErAli_LinEr = WMsgzAliLin(AyeT1Ay(Ali, "Left Right Center"))
End Property
Private Property Get WAny_Tot() As Boolean
Dim Lc As ListColumn
'For Each Lc In A_Lo.ListColumns
    'If LcFmtSpecLy_WAny_Tot(Lc, FmtSpecLy) Then WAny_Tot = True: Exit Function
'Next
End Property
Private Property Get ErAli() As String()
'ErAli = Sy(WAli_LeftRightCenter)
End Property
Private Property Get ErBdr() As String()
ErBdr = Sy(ErBdrExcessFld, ErBdrExcessLin, ErBdrDup, ErBdrFld)
End Property
Private Function ErBdr1(X$) As String()
'Return FldAy from Bdr & X
'Dim FldssAy$(): FldssAy = SSSyzAy(AywRmvT1(Bdr, X))
End Function
Private Property Get ErBdrDup() As String()
'ErBdrDup = WMsgzDup(AyDupT1(Bdr), Bdr)
End Property
Private Property Get ErBdrExcessFld() As String()
Dim LFny$(), RfNy$(), CFny$()
LFny = ErBdr1("Left")
RfNy = ErBdr1("Right")
CFny = ErBdr1("Center")
'PushIAy ErBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, LFny), "Center", "Left")
'PushIAy ErBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, RfNy), "Center", "Right")
'PushIAy ErBdrExcessFld, FmtQQ(M_Dup, AyMinus(LFny, RfNy), "Left", "Right")
End Property
Private Property Get ErBdrExcessLin() As String()
Dim L
'For Each L In Itr(AyeT1Ay(Bdr, "Left Right Center"))
'    PushI ErBdrExcessLin, FmtQQ(M_Bdr_ExcessLin, L)
'Next
End Property
Private Property Get ErBdrFld() As String()
Dim Fny$(): Fny = Sy(ErBdr1("Left"), ErBdr1("Right"), ErBdr1("Center"))
ErBdrFld = WMsgzFny(Fny, "Bdr")
End Property
Private Property Get ErBet() As String()
ErBet = SyAddAp(ErBetDup, ErBetFny, ErBetTermCnt)
End Property
Private Property Get ErBetDup() As String()
'ErBetDup = WMsgzDup(AyDupT1(Bet), Bet)
End Property
Private Property Get ErBetFny() As String()
'C$ is the col-c of Bet-line.  It should have 2 item and in Fny
'Return Er of M_Bet_* if any
End Property
Private Property Get ErBetTermCnt() As String()
Dim L
'For Each L In Itr(Bet)
    If Si(SySsl(L)) <> 3 Then
        PushI ErBetTermCnt, WMsgzBetTermCnt(L, 3)
    End If
'Next
End Property
Private Property Get ErCor() As String()
Dim L$()
'L = Cor
ErCor = SyAddAp(ErCorDup(L), ErCorFld(L), ErCorVal(L))
'Cor = L
End Property
Private Function ErCorDup(IO$()) As String()

End Function
Private Function ErCorFld(IO$()) As String()

End Function
Private Function ErCorVal(IO$()) As String()
Dim Msg$(), Er$(), L
For Each L In IO
    PushI Msg, ErCorVal1(L)
Next
'IO = AywNoEr(IO, Msg, Er)
End Function
Private Function ErCorVal1$(L)
Dim Cor$
Cor = T1(L)
End Function
Private Property Get ErFld() As String()

End Property
Private Property Get ErFldss() As String()

End Property
Private Property Get ErFldSngzDup() As String() 'It is for [SngFldLin] only.  That means T2 of LofLin is field name.  Return error msg for any FldNm is dup.
Dim T1
For Each T1 In SySsl(LofT1nnzSng) 'It is for [SngFldLin] only
    PushIAy ErFldSngzDup, ErFldSngzDup__WithinT1(T1)
Next
End Property
Private Function ErFldSngzDup__DupFld_is_fnd(DupFld, LnxAy() As Lnx, T1) As String() '[DupFld] is found within [LnxAy].  All [LnxAy] has [T1]
Dim LnoAy&(): LnoAy = LngAyzOyPrp(LnxAywT2(LnxAy, DupFld), "Lno")
Dim J%, Lno0&
For J = 1 To UB(LnoAy)
    Lno0 = LnoAy(0)
    PushI ErFldSngzDup__DupFld_is_fnd, MsgOf_Fld_Dup(LnoAy(J), T1, DupFld, Lno0)
Next
End Function
Private Function ErFldSngzDup__WithinT1(T1) As String() 'Within T1 any fld is dup?
Dim DupFld, LnxAy() As Lnx
LnxAy = WLnxAyzT1(T1)
For Each DupFld In Itr(DupT2AyzLnxAy(LnxAy))
    PushIAy ErFldSngzDup__WithinT1, ErFldSngzDup__DupFld_is_fnd(DupFld, LnxAy, T1)
Next
End Function
Private Property Get ErFml() As String()
ErFml = ErFml__InsideFmlHasInvalidFld
End Property
Private Property Get ErFml__InsideFmlHasInvalidFld() As String()
Dim I, A$, Fld$, Fml$, O$(), T1
For Each I In Itr(WLnxAyzT1("Fml"))
    With CvLnx(I)
        Asg2TRst .Lin, A, Fld, Fml
        If FstChr(Fml) <> "=" Then
            'PushI O, WMsg_Fml_FstChr(.Lno)
        Else
            Dim ErFny$(): ErFny = ErFnyzFml(Fld, Fml, A_Fny)
'            PushIAy O, ErFml__InsideFmlHasInvalidFld1(ErFny, .Lno, Fld, Fml)
        End If
    End With
Next
ErFml__InsideFmlHasInvalidFld = O
End Property
Function ErFnyzFml(Fld$, Fml$, Fny$()) As String() 'Return Subset-Fny (quote by []) in [Fml] which is error. _
It is error if any-FmlFny not in [Fny] or =[Fld]
Dim Ny$(): Ny = NyzMacro(Fml, OpnBkt:="[")
If HasEle(Ny, Fld) Then PushI ErFnyzFml, Fld
PushIAy ErFnyzFml, AyMinus(Fml, Fny)
End Function
Private Property Get ErFmt() As String()

End Property
Private Property Get ErLbl() As String()

End Property
Private Property Get ErLoNm() As String()
ErLoNm = Sy()
'1Sy(WAli_LeftRightCenter)
End Property
Private Property Get ErLvl() As String()

End Property
Private Function ErMisFnyzFmti(Fmti) As String()
'LnxAyzT1 (Fmti)
End Function
Private Property Get ErTit() As String()

End Property
Private Property Get ErTot() As String()
Dim L
'For Each L In Itr(Tot)
'    A = Avg(J)
'    Ix = IxAy(Sum, A)
'    If Ix >= 0 Then
'        Msg = FmtQQ(M_Fld_IsAvg_FndInSum, AvgLxAy(J), Avg(J), SumLxAy(Ix))
'        O.PushMsg Msg
'    End If
'Next
End Property
Private Property Get ErTot_1() '(Cnt$(), CntLxAy%(), Sum$(), SumLxAy%(), Avg$(), AvgLxAy%()) As Er
'Dim O As New Er
'Dim J%, C$, Ix%, Msg$
'For J = 0 To UB(Cnt)
'    C = Cnt(J)
'    Ix = IxAy(Sum, C)
'    If Ix >= 0 Then
'        Msg = FmtQQ(M_Fld_IsCnt_FndInSum, CntLxAy(J), Cnt(J), SumLxAy(Ix))
'        O.PushMsg Msg
'    Else
'        Ix = IxAy(Avg, C)
'        If Ix >= 0 Then
'            Msg = FmtQQ(M_Fld_IsCnt_FndInAvg, CntLxAy(J), Cnt(J), AvgLxAy(Ix))
'            O.PushMsg Msg
'        End If
'    End If
'Next
'Set1Lc ErTot_1 = O
End Property
Private Property Get ErVal() As String() 'W-Error-of-LofLinVal:W means working-value. _
which is using the some Module-Lvl-variables and it is private. _
Val here means the LofValFld of LofLin
ErVal = SyAddAp(ErValzNotNum, ErValzNotInLis, ErValzFml, ErValzNotBet)
End Property
Private Function ErValzFml() As String()

End Function
Private Function ErValzNotBet() As String()
PushIAy ErValzNotBet, ErValzNotBetz("Wdt", 10, 200)
PushIAy ErValzNotBet, ErValzNotBetz("Lvl", 2, 9)
End Function
Private Function ErValzNotBetz(T1, FmNumVal, ToNumval) As String()
'Dim Lnx(): Lnx = A_T1ToLnxAyDic(T1)
End Function
Private Function ErValzNotInLis() As String()

End Function
Private Function ErValzNotNum() As String()
Dim T
For Each T In SySsl("Wdt Lvl")
Next
End Function
Private Property Get ErWdt() As String()
End Property
Private Function WLnxAyzT1(T1) As Lnx()
If A_T1ToLnxAyDic.Exists(T1) Then WLnxAyzT1 = A_T1ToLnxAyDic(T1)
End Function
Private Function WMsgzBetTermCnt$(L, NTerm%)

End Function
Private Function WMsgzDupNy(DupNy$(), LnoStrAy$()) As String()
Dim N, J&
For Each N In Itr(DupNy)
'    PushIAy WMsgzDupNy, FmtQQ(M_Dup, N, LnoStrAy(J))
    J = J + 1
Next
End Function
Private Function WMsgzFny(Fny$(), Lin_Ty$) As String()
'Return Msg if given-Fny has some field not in A_Fny
Dim ErFny$(): ErFny = AyMinus(Fny, A_Fny)
If Si(ErFny) = 0 Then Exit Function
'PushI WMsgzFny, FmtQQ(M_Fny, ErFny, Lin_Ty)
End Function
Private Sub Z_ErBet()
'---------------
A_Fny = SySsl("A B")
'Erase Bet
'    PushI Bet, "A B C"
'    PushI Bet, "A B C"
Ept = EmpSy
'    PushIAy Ept, WMsgzDup(Sy("A"), Bet)
GoSub Tst
Exit Sub
'---------------
Tst:
    Act = ErBet
    C
    Return
End Sub
Private Sub Z_ErFldSngzDup()
Dim Lof$(), Fny$(), Act$(), Ept$()
GoSub T1
Exit Sub
T1:
    Lof = SplitVbar("Fml AA sdlkfsdflk|Fml AA skldf|Fml BB sdklfjdlf|Fml BB sdlfkjsdf|Fml BB sdklfjsdf|Fml CC sdfsdf")
    GoTo Tst
Tst:
    Init Lof, Fny
    Act = ErFldSngzDup
End Sub
Private Function MsgOf_Val_NotNum(Lno&, T1$, Val$) As String():                                             MsgOf_Val_NotNum = FmtMacro(M_Val_NotNum, Lno, T1, Val):                                          End Function
Private Function MsgOf_Val_NotBet(Lno&, T1$, Val$, FmNo) As String():                                       MsgOf_Val_NotBet = FmtMacro(M_Val_NotBet, Lno, T1, Val, FmNo):                                    End Function
Private Function MsgOf_Val_NotInLis(Lno&, T1$, ErVal$, VdtValNm$) As String():                              MsgOf_Val_NotInLis = FmtMacro(M_Val_NotInLis, Lno, T1, ErVal, VdtValNm):                          End Function
Private Function MsgOf_Val_FmlFld(Lno&, Fml$, ErFny$(), VdtFny$()) As String():                             MsgOf_Val_FmlFld = FmtMacro(M_Val_FmlFld, Lno, Fml, ErFny, VdtFny):                               End Function
Private Function MsgOf_Val_FmlNotBegEq(Lno&, Fml$) As String():                                             MsgOf_Val_FmlNotBegEq = FmtMacro(M_Val_FmlNotBegEq, Lno, Fml):                                    End Function
Private Function MsgOf_Fld_NotInFny(Lno&, T1$, F) As String():                                              MsgOf_Fld_NotInFny = FmtMacro(M_Fld_NotInFny, Lno, T1, F):                                        End Function
Private Function MsgOf_Fld_Dup(Lno&, T1, F, AlreadyInLno&) As String():                                     MsgOf_Fld_Dup = FmtMacro(M_Fld_Dup, Lno, T1, F, AlreadyInLno):                                    End Function
Private Function MsgOf_Fldss_NotSel(Lno&, T1$, Fldss$) As String():                                         MsgOf_Fldss_NotSel = FmtMacro(M_Fldss_NotSel, Lno, T1, Fldss):                                    End Function
Private Function MsgOf_Fldss_DupSel(Lno&, T1$) As String():                                                 MsgOf_Fldss_DupSel = FmtMacro(M_Fldss_DupSel, Lno, T1):                                           End Function
Private Function MsgOf_LoNm(Lno&, Val$) As String():                                                        MsgOf_LoNm = FmtMacro(M_LoNm, Lno, Val):                                                          End Function
Private Function MsgOf_LoNm_Mis() As String():                                                              MsgOf_LoNm_Mis = FmtMacro(M_LoNm_Mis):                                                            End Function
Private Function MsgOf_LoNm_Dup(Lno&, AlreadyInLno&) As String():                                           MsgOf_LoNm_Dup = FmtMacro(M_LoNm_Dup, Lno, AlreadyInLno):                                         End Function
Private Function MsgOf_Tot_DupSel(Lno&, TotKd$, Fldss$, SelFld$, AlreadyInLno&, AlreadyTotKd$) As String(): MsgOf_Tot_DupSel = FmtMacro(M_Tot_DupSel, Lno, TotKd, Fldss, SelFld, AlreadyInLno, AlreadyTotKd): End Function
Private Function MsgOf_Bet_3Fld(Lno&) As String():                                                          MsgOf_Bet_3Fld = FmtMacro(M_Bet_3Fld, Lno):                                                       End Function
Private Function MsgOf_Bet_EqFmTo(Lno&) As String():                                                        MsgOf_Bet_EqFmTo = FmtMacro(M_Bet_EqFmTo, Lno):                                                   End Function
Private Function MsgOf_Bet_FldSeq(Lno&) As String():                                                        MsgOf_Bet_FldSeq = FmtMacro(M_Bet_FldSeq, Lno):                                                   End Function

