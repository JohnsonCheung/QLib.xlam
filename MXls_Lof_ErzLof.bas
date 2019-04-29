Attribute VB_Name = "MXls_Lof_ErzLof"
Option Explicit
Const CMod$ = ""
Public Const LofT1nn$ = _
                            "Lo Ali Bdr Tot Wdt Fmt Lvl Cor Fml Lbl Tit Bet" ' Fmt. i.tm s.pace s.eparated string
Public Const LofT1nnzSng$ = "                               Fml Lbl Tit Bet" ' Sng.sigle field per line
Public Const LofT1nnzMul$ = "Lo Ali Bdr Tot Wdt Fmt Lvl Cor                " ' Mul.tiple field per line
'GenErzMsg-Src-Beg.
'Val_NotNum      Lno#{Lno&} is [{T1$}] line having Val({Val$}) which should be a number
'Val_NotBet      Lno#{Lno&} is [{T1$}] line having Val({Val$}) which between ({FmNo}) and (ToNm})
'Val_NotInLis    Lno#{Lno&} is [{T1$}] line having invalid Val({ErzVal$}).  See valid-value-{VdtValNm$}
'Val_FmlFld      Lno#{Lno&} is [Fml] line having invalid Fml({Fml$}) due to invalid Fny({ErzFny$()}).  Valid-Fny are [{VdtFny$()}]
'Val_FmlNotBegEq Lno#{Lno&} is [Fml] line having [{Fml$}] which is not started with [=]
'Fld_NotInFny    Lno#{Lno&} is [{T1$}] line having Fld({F}) which should one of the Fny value.  See [Fny-Value]
'Fld_Dup         Lno#{Lno&} is [{T1}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno&}
'Fldss_NotSel    Lno#{Lno&} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]
'Fldss_DupSel    Lno#{Lno&} is [{T1$}] line having
'LoNm            Lno#{Lno&} is [Lo-Nm] line having value({Val$}) which is not a good name.
'LoNm_Mis        [Lo-Nm] line is missing
'LoNm_Dup        Lno#{Lno&} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno&}
'Tot_DupSel      Lno#{Lno&} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno&} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored.
'Bet_N3Fld        Lno#{Lno&} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]"
'Bet_EqFmTo      Lno#{Lno&} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal.
'Bet_FldSeq      Lno#{Lno&} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]"
'GenErzMsg-Src-End.
Private A$(), A_Fny$(), A_T1ToLnxAyDic As Dictionary
Private Const M_Val_NotNum$ = "Lno#{Lno&} is [{T1$}] line having Val({Val$}) which should be a number"
Private Const M_Val_NotBet$ = "Lno#{Lno&} is [{T1$}] line having Val({Val$}) which between ({FmNo}) and (ToNm})"
Private Const M_Val_NotInLis$ = "Lno#{Lno&} is [{T1$}] line having invalid Val({ErzVal$}).  See valid-value-{VdtValNm$}"
Private Const M_Val_FmlFld$ = "Lno#{Lno&} is [Fml] line having invalid Fml({Fml$}) due to invalid Fny({ErzFny$()}).  Valid-Fny are [{VdtFny$()}]"
Private Const M_Val_FmlNotBegEq$ = "Lno#{Lno&} is [Fml] line having [{Fml$}] which is not started with [=]"
Private Const M_Fld_NotInFny$ = "Lno#{Lno&} is [{T1$}] line having Fld({F}) which should one of the Fny value.  See [Fny-Value]"
Private Const M_Fld_Dup$ = "Lno#{Lno&} is [{T1}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno&}"
Private Const M_Fldss_NotSel$ = "Lno#{Lno&} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]"
Private Const M_Fldss_DupSel$ = "Lno#{Lno&} is [{T1$}] line having"
Private Const M_LoNm$ = "Lno#{Lno&} is [Lo-Nm] line having value({Val$}) which is not a good name."
Private Const M_LoNm_Mis$ = "[Lo-Nm] line is missing"
Private Const M_LoNm_Dup$ = "Lno#{Lno&} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno&}"
Private Const M_Tot_DupSel$ = "Lno#{Lno&} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno&} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored."
Private Const M_Bet_N3Fld$ = "Lno#{Lno&} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]"""
Private Const M_Bet_EqFmTo$ = "Lno#{Lno&} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal."
Private Const M_Bet_FldSeq$ = "Lno#{Lno&} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]"""
Private Function MsgVal_FmlNotBegWithEq$(Lno&, Fml$)

End Function
Property Get LofT1Ny() As String()
LofT1Ny = TermAy(LofT1nn)
End Property
Function ErzLof(Lof$(), Fny$()) As String() 'Erzror-of-ListObj-Formatter:Erz.z.Lo.f
Const CSub$ = CMod & "LofErz"
Init Lof, Fny
ErzLof = SyAddAp( _
    ErzVal, ErzFld, ErzFldss, ErzLoNm, _
    ErzAli, ErzBdr, ErzTot, _
    ErzWdt, ErzFmt, ErzLvl, ErzCor, _
    ErzFml, ErzLbl, ErzTit, ErzBet)
End Function
Function FnywLikssAy(Fny$(), LikssAy$()) As String()
Dim F$, I, LikeAy$()
LikeAy = TermAsetzTLinAy(LikssAy).Sy
For Each I In Itr(Fny)
    F = I
    If HitLikAy(F, LikeAy) Then PushI FnywLikssAy, F
Next
End Function
Private Sub Init(Lof$(), Fny$())
A = Lof
A_Fny = Fny
Set A_T1ToLnxAyDic = LnxAyDiczT1nn(A, LofT1nn)
End Sub
Private Property Get WAli_LeftRightCenter() As String()
'ErzAli_LinErz = WMsgzAliLin(AyeT1Ay(Ali, "Left Right Center"))
End Property
Private Property Get WAny_Tot() As Boolean
Dim Lc As ListColumn
'For Each Lc In A_Lo.ListColumns
    'If LcFmtSpecLy_WAny_Tot(Lc, FmtSpecLy) Then WAny_Tot = True: Exit Function
'Next
End Property
Private Property Get ErzAli() As String()
'ErzAli = Sy(WAli_LeftRightCenter)
End Property
Private Property Get ErzBdr() As String()
ErzBdr = Sy(ErzBdrExcessFld, ErzBdrExcessLin, ErzBdrDup, ErzBdrFld)
End Property
Private Function ErzBdr1(X$) As String()
'Return FldAy from Bdr & X
'Dim FldssAy$(): FldssAy = SSSyzAy(AywRmvT1(Bdr, X))
End Function
Private Property Get ErzBdrDup() As String()
'ErzBdrDup = WMsgzDup(DupT1(Bdr), Bdr)
End Property
Private Property Get ErzBdrExcessFld() As String()
Dim LFny$(), RfNy$(), CFny$()
LFny = ErzBdr1("Left")
RfNy = ErzBdr1("Right")
CFny = ErzBdr1("Center")
'PushIAy ErzBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, LFny), "Center", "Left")
'PushIAy ErzBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, RfNy), "Center", "Right")
'PushIAy ErzBdrExcessFld, FmtQQ(M_Dup, AyMinus(LFny, RfNy), "Left", "Right")
End Property
Private Property Get ErzBdrExcessLin() As String()
Dim L
'For Each L In Itr(AyeT1Ay(Bdr, "Left Right Center"))
'    PushI ErzBdrExcessLin, FmtQQ(M_Bdr_ExcessLin, L)
'Next
End Property
Private Property Get ErzBdrFld() As String()
Dim Fny$(): Fny = Sy(ErzBdr1("Left"), ErzBdr1("Right"), ErzBdr1("Center"))
ErzBdrFld = WMsgzFny(Fny, "Bdr")
End Property
Private Property Get ErzBet() As String()
ErzBet = SyAddAp(ErzBetDup, ErzBetFny, ErzBetTermCnt)
End Property
Private Property Get ErzBetDup() As String()
'ErzBetDup = WMsgzDup(DupT1(Bet), Bet)
End Property
Private Property Get ErzBetFny() As String()
'C$ is the col-c of Bet-line.  It should have 2 item and in Fny
'Return Erz of M_Bet_* if any
End Property
Private Property Get ErzBetTermCnt() As String()
Dim L$, I
'For Each L In Itr(Bet)
    L = I
    If Si(SySsl(L)) <> 3 Then
        PushI ErzBetTermCnt, WMsgzBetTermCnt(L, 3)
    End If
'Next
End Property
Private Property Get ErzCor() As String()
Dim L$()
'L = Cor
ErzCor = SyAddAp(ErzCorDup(L), ErzCorFld(L), ErzCorVal(L))
'Cor = L
End Property
Private Function ErzCorDup(IO$()) As String()

End Function
Private Function ErzCorFld(IO$()) As String()

End Function
Private Function ErzCorVal(IO$()) As String()
Dim Msg$(), Erz$(), L$, I
For Each I In IO
    L = I
    PushI Msg, ErzCorVal1(L)
Next
'IO = AywNoErz(IO, Msg, Erz)
End Function
Private Function ErzCorVal1$(L$)
Dim Cor$
Cor = T1(L)
End Function
Private Property Get ErzFld() As String()

End Property
Private Property Get ErzFldss() As String()

End Property
Private Property Get ErzFldSngzDup() As String() 'It is for [SngFldLin] only.  That means T2 of LofLin is field name.  Return error msg for any FldNm is dup.
Dim T1$, I
For Each I In SySsl(LofT1nnzSng) 'It is for [SngFldLin] only
    T1 = I
    PushIAy ErzFldSngzDup, ErzFldSngzDup__WithinT1(T1)
Next
End Property
Private Function ErzFldSngzDup__DupFld_is_fnd(DupFld$, LnxAy() As Lnx, T1$) As String() '[DupFld] is found within [LnxAy].  All [LnxAy] has [T1]
Dim LnoAy&(): LnoAy = LngAyzOyPrp(LnxAywT2(LnxAy, DupFld), "Lno")
Dim J%, Lno0&
For J = 1 To UB(LnoAy)
    Lno0 = LnoAy(0)
    PushI ErzFldSngzDup__DupFld_is_fnd, MsgOf_Fld_Dup(LnoAy(J), T1, DupFld, Lno0)
Next
End Function
Private Function ErzFldSngzDup__WithinT1(T1$) As String() 'Within T1 any fld is dup?
Dim DupFld$, I, LnxAy() As Lnx
LnxAy = WLnxAyzT1(T1)
For Each I In Itr(DupT2AyzLnxAy(LnxAy))
    DupFld = I
    PushIAy ErzFldSngzDup__WithinT1, ErzFldSngzDup__DupFld_is_fnd(DupFld, LnxAy, T1)
Next
End Function
Private Property Get ErzFml() As String()
ErzFml = ErzFml__InsideFmlHasInvalidFld
End Property
Private Property Get ErzFml__InsideFmlHasInvalidFld() As String()
Dim I, A$, Fld$, Fml$, O$(), T1
For Each I In Itr(WLnxAyzT1("Fml"))
    With CvLnx(I)
        AsgN2tRst .Lin, A, Fld, Fml
        If FstChr(Fml) <> "=" Then
            'PushI O, WMsg_Fml_FstChr(.Lno)
        Else
            Dim ErzFny$(): ErzFny = ErzFnyzFml(Fld, Fml, A_Fny)
'            PushIAy O, ErzFml__InsideFmlHasInvalidFld1(ErzFny, .Lno, Fld, Fml)
        End If
    End With
Next
ErzFml__InsideFmlHasInvalidFld = O
End Property
Function ErzFnyzFml(Fld$, Fml$, Fny$()) As String() 'Return Subset-Fny (quote by []) in [Fml] which is error. _
It is error if any-FmlFny not in [Fny] or =[Fld]
Dim Ny$(): Ny = NyzMacro(Fml, OpnBkt:="[")
If HasEle(Ny, Fld) Then PushI ErzFnyzFml, Fld
PushIAy ErzFnyzFml, AyMinus(Fml, Fny)
End Function
Private Property Get ErzFmt() As String()

End Property
Private Property Get ErzLbl() As String()

End Property
Private Property Get ErzLoNm() As String()
ErzLoNm = Sy()
'1Sy(WAli_LeftRightCenter)
End Property
Private Property Get ErzLvl() As String()

End Property
Private Function ErzMisFnyzFmti(Fmti) As String()
'LnxAyzT1 (Fmti)
End Function
Private Property Get ErzTit() As String()

End Property
Private Property Get ErzTot() As String()
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
Private Property Get ErzTot_1() '(Cnt$(), CntLxAy%(), Sum$(), SumLxAy%(), Avg$(), AvgLxAy%()) As Erz
'Dim O As New Erz
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
'Set1Lc ErzTot_1 = O
End Property
Private Property Get ErzVal() As String() 'W-Erzror-of-LofLinVal:W means working-value. _
which is using the some Module-Lvl-variables and it is private. _
Val here means the LofValFld of LofLin
ErzVal = SyAddAp(ErzValOfNotNum, ErzValOfNotInLis, ErzValOfFml, ErzValOfNotBet)
End Property
Private Function ErzValOfFml() As String()

End Function
Private Function ErzValOfNotBet() As String()
PushIAy ErzValOfNotBet, ErzValOfNotBetz("Wdt", 10, 200)
PushIAy ErzValOfNotBet, ErzValOfNotBetz("Lvl", 2, 9)
End Function
Private Function ErzValOfNotBetz(T1, FmNumVal, ToNumval) As String()
'Dim Lnx(): Lnx = A_T1ToLnxAyDic(T1)
End Function
Private Function ErzValOfNotInLis() As String()

End Function
Private Function ErzValOfNotNum() As String()
Dim T
For Each T In SySsl("Wdt Lvl")
Next
End Function
Private Property Get ErzWdt() As String()
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
Dim ErzFny$(): ErzFny = AyMinus(Fny, A_Fny)
If Si(ErzFny) = 0 Then Exit Function
'PushI WMsgzFny, FmtQQ(M_Fny, ErzFny, Lin_Ty)
End Function
Private Sub Z_ErzBet()
'---------------
A_Fny = SySsl("A B")
'Erzase Bet
'    PushI Bet, "A B C"
'    PushI Bet, "A B C"
Ept = EmpSy
'    PushIAy Ept, WMsgzDup(Sy("A"), Bet)
GoSub Tst
Exit Sub
'---------------
Tst:
    Act = ErzBet
    C
    Return
End Sub
Private Sub Z_ErzFldSngzDup()
Dim Lof$(), Fny$(), Act$(), Ept$()
GoSub T1
Exit Sub
T1:
    Lof = SplitVBar("Fml AA sdlkfsdflk|Fml AA skldf|Fml BB sdklfjdlf|Fml BB sdlfkjsdf|Fml BB sdklfjsdf|Fml CC sdfsdf")
    GoTo Tst
Tst:
    Init Lof, Fny
    Act = ErzFldSngzDup
End Sub
Private Function MsgOf_Val_NotNum(Lno&, T1$, Val$) As String():                                             MsgOf_Val_NotNum = FmtMacro(M_Val_NotNum, Lno, T1, Val):                                          End Function
Private Function MsgOf_Val_NotBet(Lno&, T1$, Val$, FmNo) As String():                                       MsgOf_Val_NotBet = FmtMacro(M_Val_NotBet, Lno, T1, Val, FmNo):                                    End Function
Private Function MsgOf_Val_NotInLis(Lno&, T1$, ErzVal$, VdtValNm$) As String():                             MsgOf_Val_NotInLis = FmtMacro(M_Val_NotInLis, Lno, T1, ErzVal, VdtValNm):                          End Function
Private Function MsgOf_Val_FmlFld(Lno&, Fml$, ErzFny$(), VdtFny$()) As String():                            MsgOf_Val_FmlFld = FmtMacro(M_Val_FmlFld, Lno, Fml, ErzFny, VdtFny):                               End Function
Private Function MsgOf_Val_FmlNotBegEq(Lno&, Fml$) As String():                                             MsgOf_Val_FmlNotBegEq = FmtMacro(M_Val_FmlNotBegEq, Lno, Fml):                                    End Function
Private Function MsgOf_Fld_NotInFny(Lno&, T1$, F) As String():                                              MsgOf_Fld_NotInFny = FmtMacro(M_Fld_NotInFny, Lno, T1, F):                                        End Function
Private Function MsgOf_Fld_Dup(Lno&, T1, F, AlreadyInLno&) As String():                                     MsgOf_Fld_Dup = FmtMacro(M_Fld_Dup, Lno, T1, F, AlreadyInLno):                                    End Function
Private Function MsgOf_Fldss_NotSel(Lno&, T1$, Fldss$) As String():                                         MsgOf_Fldss_NotSel = FmtMacro(M_Fldss_NotSel, Lno, T1, Fldss):                                    End Function
Private Function MsgOf_Fldss_DupSel(Lno&, T1$) As String():                                                 MsgOf_Fldss_DupSel = FmtMacro(M_Fldss_DupSel, Lno, T1):                                           End Function
Private Function MsgOf_LoNm(Lno&, Val$) As String():                                                        MsgOf_LoNm = FmtMacro(M_LoNm, Lno, Val):                                                          End Function
Private Function MsgOf_LoNm_Mis() As String():                                                              MsgOf_LoNm_Mis = FmtMacro(M_LoNm_Mis):                                                            End Function
Private Function MsgOf_LoNm_Dup(Lno&, AlreadyInLno&) As String():                                           MsgOf_LoNm_Dup = FmtMacro(M_LoNm_Dup, Lno, AlreadyInLno):                                         End Function
Private Function MsgOf_Tot_DupSel(Lno&, TotKd$, Fldss$, SelFld$, AlreadyInLno&, AlreadyTotKd$) As String(): MsgOf_Tot_DupSel = FmtMacro(M_Tot_DupSel, Lno, TotKd, Fldss, SelFld, AlreadyInLno, AlreadyTotKd): End Function
Private Function MsgOf_Bet_N3Fld(Lno&) As String():                                                         MsgOf_Bet_N3Fld = FmtMacro(M_Bet_N3Fld, Lno):                                                       End Function
Private Function MsgOf_Bet_EqFmTo(Lno&) As String():                                                        MsgOf_Bet_EqFmTo = FmtMacro(M_Bet_EqFmTo, Lno):                                                   End Function
Private Function MsgOf_Bet_FldSeq(Lno&) As String():                                                        MsgOf_Bet_FldSeq = FmtMacro(M_Bet_FldSeq, Lno):                                                   End Function

