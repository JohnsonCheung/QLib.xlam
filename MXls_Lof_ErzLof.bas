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
'Fld_Dup         Lno#{Lno&} is [{T1$}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno&}
'Fldss_NotSel    Lno#{Lno&} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]
'Fldss_DupSel    Lno#{Lno&} is [{T1$}] line having
'LoNm            Lno#{Lno&} is [Lo-Nm] line having value({Val$}) which is not a good name"
'LoNm_Mis        [Lo-Nm] line is missing
'LoNm_Dup        Lno#{Lno&} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno&}
'Tot_DupSel      Lno#{Lno&} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno&} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored.
'Bet_3Fld        Lno#{Lno&} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]"
'Bet_EqFmTo      Lno#{Lno&} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal."
'Bet_FldSeq      Lno#{Lno&} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]"
'GenErMsg-Src-End.
Const M01$ = "Lno#? is [?] line having Val(?) which should be a number" 'MVal_IsNotNum
Const M02$ = "Lno#? is [?] line having Val(?) which between (?) and (?)" 'MVal_IsNotBet
Const M03$ = "Lno#? is [?] line having invalid Val(?).  See valid-value-?"  'For Ali Bdr Tot Cor
Const M04$ = "Lno#? is [Fml] line having invalid Fml(?) due to invalid Fny{?}.  Valid-Fny are [?]." 'For Fml
Const M04a$ = "Lno#? is [Fml] line having invalid Fml(?) due to first char is not [=]"
Const M05$ = "Lno#? is [?] line having Fld(?) which should one of the Fny value.  See [Fny-Value]" 'For Fml Lbl Tit Bet
Const M06$ = "Lno#? is [?] line having Fld(?) which is duplicated and ignored due to it has defined in Lno#?" 'For
Const M07$ = "Lno#? is [?] line having Fldss(?) which should select one for Fny value.  See [Fny-Value]"
Const M08$ = "Lno#? is LoNm line having value(?) which is not a good name"
Const M09$ = "LoNm line is missing"
Const M10$ = "Lno#? is LoNm line which is duplicated and ignored due to there is already a LoNm in Lno#?"
Const M11$ = "Lno#? is [Bdr ?] line having Fld(?) which is duplicated and ignored due to it is alredy defined in Lno#? as [?]"
Const M12$ = "Lno#? is [Bdr ?] line which is duplicated and ignore due to there is already a [Bdr ?] line in Lno#?"
Const M13$ = "Lno#? is [Tot ?] line. It has selected more than one fields: [?].  [?] is used, the rest is ignored"
Const M14$ = "Lno#? is [] line having Val(?) which in invalid.  Valid value is one of (Tot Avg Sum)"
Const M15$ = "Lno#? is [?] line with Fld(?), which is already defined as (?)-Fld in same line and is ignored."
Const M16$ = "Lno#? is [?] line with Fld(?), which is already defined as (?)-Fld in Lno#? and is ignored"
Const M17$ = "Lno#? is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]"
Const M18$ = "Lno#? is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal."
Const M19$ = "Lno#? is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]"
Const MVal_IsNotNum$ = M01
Const MVal_IsNotBet$ = M02
Const MVal_IsNotInLis$ = M03
Const MVal_Fml$ = M04
Const MVal_FmlNotBegWithEq$ = M04a
Const MFld_Mis$ = M05
Const MFld_Dup$ = M06
Const MFldss_NotSelAnyFld$ = M07
Const MLoNm_NoVal$ = M08
Const MLoNm_NoLin$ = M09
Const MLoNm_Excess$ = M10
Const MBdr_ExcessFld$ = M11
Const MBdr_ExcessLin$ = M12
Const MTot_DupFld$ = M13
Const MTot_MoreThanOneFldIsCnt$ = M14
Const MTot_Fld_AlreadyDefinedSamLin$ = M15
Const MTot_Fld_AlreadyDefinedInLno$ = M16
Const MBet_TermCntEr = M17
Const MBet_FmToEq = M18
Const MBet_Bet = M19
Private A$(), A_Fny$(), A_T1ToLnxAyDic As Dictionary
Private Function MsgVal_FmlNotBegWithEq$(Lno&, Fml$)

End Function
Function ErzLof(Lof$(), Fny$()) As String() 'Error-of-ListObj-Formatter:Er.z.Lo.f
Const CSub$ = CMod & "LofEr"
Init Lof, Fny
ErzLof = AyAddAp( _
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

Private Function WMsg_Fld_Dup$(Lno&, T1, DupFld, AlreadInLno&)
WMsg_Fld_Dup = FmtQQ(MFld_Dup, Lno, T1, DupFld, AlreadInLno)
End Function

Private Property Get ErAli() As String()
ErAli = Sy(WAli_LeftRightCenter)
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
ErBet = Sy(ErBetDup, ErBetFny, ErBetTermCnt)
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
    If Sz(SySsl(L)) <> 3 Then
        PushI ErBetTermCnt, WMsgzBetTermCnt(L, 3)
    End If
'Next
End Property

Private Property Get ErCor() As String()
Dim L$()
'L = Cor
ErCor = Sy(ErCorDup(L), ErCorFld(L), ErCorVal(L))
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
    PushI ErFldSngzDup__DupFld_is_fnd, WMsg_Fld_Dup(LnoAy(J), T1, DupFld, Lno0)
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
ErFml = Sy(ErFml__InsideFmlHasInvalidFld)
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
Dim Lnx(): Lnx = A_T1ToLnxAyDic(T1)
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
If Sz(ErFny) = 0 Then Exit Function
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
    Lof = SplitVBar("Fml AA sdlkfsdflk|Fml AA skldf|Fml BB sdklfjdlf|Fml BB sdlfkjsdf|Fml BB sdklfjsdf|Fml CC sdfsdf")
    GoTo Tst
Tst:
    Init Lof, Fny
    Act = ErFldSngzDup
   
End Sub

