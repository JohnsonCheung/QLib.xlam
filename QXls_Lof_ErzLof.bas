Attribute VB_Name = "QXls_Lof_ErzLof"
Option Compare Text
Option Explicit
Private Const Asm$ = "QXls"
Private Const CMod$ = "MXls_Lof_ErzLof."
Public Const LofT1nn$ = _
                            "Lo Ali Bdr Tot Wdt Fmt Lvl Cor Fml Lbl Tit Bet" ' Fmt. i.tm s.pace s.eparated string
Public Const LofT1nnzSng$ = "                               Fml Lbl Tit Bet" ' Sng.sigle field per line
Public Const LofT1nnzMul$ = "Lo Ali Bdr Tot Wdt Fmt Lvl Cor                " ' Mul.tiple field per line
'GenErzMsg-Src-Beg.
'Val_NotNum      Lno#{Lno} is [{T1$}] line having Val({Val$}) which should be a number
'Val_NotBet      Lno#{Lno} is [{T1$}] line having Val({Val$}) which between ({FmNo}) and (ToNm})
'Val_NotInLis    Lno#{Lno} is [{T1$}] line having invalid Val({ErzVal$}).  See valid-value-{VdtValNm$}
'Val_FmlFld      Lno#{Lno} is [Fml] line having invalid Fml({Fml$}) due to invalid Fny({ErzFny$()}).  Valid-Fny are [{VdtFny$()}]
'Val_FmlNotBegEq Lno#{Lno} is [Fml] line having [{Fml$}] which is not started with [=]
'Fld_NotInFny    Lno#{Lno} is [{T1$}] line having Fld({F}) which should one of the Fny value.  See [Fny-Value]
'Fld_Dup         Lno#{Lno} is [{T1}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno}
'Fldss_NotSel    Lno#{Lno} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]
'Fldss_DupSel    Lno#{Lno} is [{T1$}] line having
'Lo_ErNm           Lno#{Lno} is [Lo-Nm] line having value({Val$}) which is not a good name.
'Lo_MisNm          [Lo-Nm] line is missing
'Lo_DupNm          Lno#{Lnoss$} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno}
'Tot_MustAvgCntSum Lno#({Lnoss$}) is [Tot] line having 2nd term value.  Valid 2nd term value is [Avg Cnt Sum]
'Tot_DupSel        Lno#{Lno} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored.
'Bet_N3Fld          Lno#{Lno} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]"
'Bet_EqFmTo         Lno#{Lno} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal.
'Bet_FldSeq         Lno#{Lno} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]"
'GenErzMsg-Src-End.
Private Const C_Tot_MustAvgCntSum$ = ""
Private Const C_Lo_MisNm$ = ""
Private Const C_Lo_ErNm = ""
Private Const MAli_MustLRCenter$ = ""

Private Function M_Lo_ErNm$(LnoAy&())
If Si(LnoAy) = 0 Then Exit Function
M_Lo_ErNm = FmtQQ(C_Lo_ErNm, JnSpc(LnoAy))
End Function

Function LofT1Ny() As String()
LofT1Ny = TermAy(LofT1nn)
End Function

Private Function B_ELo_DupNm(Dup As Lnxses) As String()

End Function

Private Function B_ELo_MisFld(IsMisLoFld As Boolean) As String()

End Function
Private Function B_ELo_DupLofFld(DupFny$()) As String()

End Function
Private Function B_ELo_IsMisNm$(IsMisNm As Boolean)
Stop
'B_ELo_IsMisNm = SzIf(IsMisNm, MLo_MisNm)
End Function

Private Function B_ELo_ErNm$(LnoAy() As Long)
B_ELo_ErNm = M_Lo_ErNm(LnoAy)
End Function
Private Function B_ELo_MisNm(IsMisNm As Boolean)

End Function
Function ErzLof(Lof$(), LoFny$()) As String() 'Erzror-of-ListObj-Formatter:Erz.z.Lo.f
Dim Lo_IsMisNm As Boolean
Dim Lo_DupNm As Lnxses
Dim Lo_ErNmLnoAy&()
Dim Lo_IsMisFldLin As Boolean
Dim Lo_DupFny$()
Dim LofFny$()
'--
Dim AliKwErLnoss$
Dim DupLofFny$()
Dim ELo1$:   ELo1 = B_ELo_MisNm(Lo_IsMisNm)
Dim ELo2$:   ELo2 = B_ELo_ErNm(Lo_ErNmLnoAy)
Dim ELo3$(): ELo3 = B_ELo_DupNm(Lo_DupNm)
Dim ELo4$(): ELo4 = B_ELo_MisFld(Lo_IsMisFldLin)
Dim ELo5$(): ELo5 = B_ELo_DupLofFld(DupLofFny)
Dim ELo6$(): ' ELo6 = B_ELo_NoLofFny(LofFny)

Dim EAli1$(): EAli1 = Sy(FmtQQ(MAli_MustLRCenter, AliKwErLnoss))
Dim EWdt1$(): EWdt1 = B_EVal_NotBet("Wdt", 10, 200)
Dim ELvl1$()
Dim EFmt1$()
Dim EFmt2$()
Dim EFmt3$()
Dim ELvl2$(): ELvl2 = B_EVal_NotBet("Lvl", 2, 9)
Dim ELvl3$()
Dim ECor1$()
Dim ECor2$()
Dim ECor3$()
Dim EFml1$()
Dim EFml2$()
Dim EFml3$()
Dim ELbl1$()
Dim ELbl2$()
Dim ELbl3$()
Dim ETit1$()
Dim ETit2$()
Dim ETit3$()
Dim EBet1$()
Dim EBet2$()
Dim EBet3$()

Dim ELo$(): ELo = Sy(ELo1, ELo2, ELo3)
Dim EAli$(): EAli = Sy(EAli1)
Dim EBdr$(): ELo = Sy(ELo1, ELo2, ELo3)
Dim ETot$(): ELo = Sy(ELo1, ELo2, ELo3)
Dim EWdt$(): EWdt = Sy(EWdt1)
Dim EFmt$(): EFmt = Sy(EFmt1, EFmt2, EFmt3)
Dim ELvl$(): ELvl = Sy(ELvl1, ELvl2, ELvl3)
Dim ECor$(): ECor = Sy(ECor1, ECor2, ECor3)
Dim EFml$(): EFml = Sy(EFml1, EFml2, EFml3)
Dim ELbl$(): ELbl = Sy(ELbl1, ELbl2, ELbl3)
Dim ETit$(): ETit = Sy(ETit1)
Dim EBet$(): EBet = Sy(EBet1)
ErzLof = Sy(ELo, EAli, EBdr, ETot, EWdt, EFmt, ELvl, ECor, EFml, ELbl, ETit, EBet)
End Function
Function FnywLikssAy(Fny$(), LikssAy$()) As String()
Dim F$, I, LikAy$()
LikAy = TermAsetzTLiny(LikssAy).Sy
For Each I In Itr(Fny)
    F = I
    If HitLikAy(F, LikAy) Then PushI FnywLikssAy, F
Next
End Function

Private Function WAli_LeftRightCenter() As String()
'ErzAli_LinErz = WMsgzAliLin(SyeT1Sy(Ali, "Left Right Center"))
End Function
Private Function WAny_Tot() As Boolean
Dim Lc As ListColumn
'For Each Lc In A_Lo.ListColumns
    'If LcFmtSpecLy_WAny_Tot(Lc, FmtSpecLy) Then WAny_Tot = True: Exit Function
'Next
End Function
Private Function ErzBdr1(X$) As String()
'Return FldAy from Bdr & X
'Dim FldssAy$(): FldssAy = SSSyzAy(AywRmvT1(Bdr, X))
End Function
Private Function B_EBdr_Dup() As String()
'ErzBdrDup = WMsgzDup(DupT1(Bdr), Bdr)
End Function
Private Function ErzBdrExcessFld() As String()
Dim LFny$(), RfNy$(), CFny$()
LFny = ErzBdr1("Left")
RfNy = ErzBdr1("Right")
CFny = ErzBdr1("Center")
'PushIAy ErzBdrExcessFld, FmtQQ(M_Dup, MinusAy(CFny, LFny), "Center", "Left")
'PushIAy ErzBdrExcessFld, FmtQQ(M_Dup, MinusAy(CFny, RfNy), "Center", "Right")
'PushIAy ErzBdrExcessFld, FmtQQ(M_Dup, MinusAy(LFny, RfNy), "Left", "Right")
End Function
Private Function ErzBdrExcessLin() As String()
Dim L
'For Each L In Itr(SyeT1Sy(Bdr, "Left Right Center"))
'    PushI ErzBdrExcessLin, FmtQQ(M_Bdr_ExcessLin, L)
'Next
End Function
Private Function ErzBdrFld() As String()
Dim Fny$(): Fny = Sy(ErzBdr1("Left"), ErzBdr1("Right"), ErzBdr1("Center"))
ErzBdrFld = WMsgzFny(Fny, "Bdr")
End Function
Private Function ErzBet() As String()
ErzBet = Sy(ErzBetDup, ErzBetFny, ErzBetTermCnt)
End Function
Private Function ErzBetDup() As String()
'ErzBetDup = WMsgzDup(DupT1(Bet), Bet)
End Function
Private Function ErzBetFny() As String()
'C$ is the col-c of Bet-line.  It should have 2 item and in Fny
'Return Erz of M_Bet_* if any
End Function
Private Function ErzBetTermCnt() As String()
Dim L$, I
'For Each L In Itr(Bet)
    L = I
    If Si(SyzSS(L)) <> 3 Then
        PushI ErzBetTermCnt, WMsgzBetTermCnt(L, 3)
    End If
'Next
End Function
Private Function ErzCor() As String()
Dim L$()
'L = Cor
ErzCor = Sy(ErzCorDup(L), ErzCorFld(L), ErzCorVal(L))
'Cor = L
End Function
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
Private Function B_EFld() As String()

End Function
Private Function ErzFldss() As String()

End Function
Private Function ErzFldSngzDup(Fny$(), Lof$()) As String() 'It is for [SngFldLin] only.  That means T2 of LofLin is field name.  Return error msg for any FldNm is dup.
Dim T1$, I
For Each I In SyzSS(LofT1nnzSng) 'It is for [SngFldLin] only
    T1 = I
    PushIAy ErzFldSngzDup, ErzFldSngzDup__WithinT1(T1)
Next
End Function
Private Function ErzFldSngzDup__DupFld_is_fnd(DupFld$, Lnxs As Lnxs, T1) As String() '[DupFld] is found within [Lnxs].  All [Lnxs] has [T1]
Dim LnoAy&(): ' LnoAy = LngAyzOyPrp(LnxswT2(Lnxs, DupFld), "Lno")
Dim J%, Lno0&
For J = 1 To UB(LnoAy)
    Lno0 = LnoAy(0)
'    PushI ErzFldSngzDup__DupFld_is_fnd, MsgOf_Fld_Dup(LnoAy(J), T1, DupFld, Lno0)
Next
End Function
Private Function ErzFldSngzDup__WithinT1(T1) As String() 'Within T1 any fld is dup?
Dim DupFld$, I, Lnxs As Lnxs
Lnxs = WLnxszT1(T1)
For Each I In Itr(DupT2AyzLnxs(Lnxs))
    DupFld = I
'    PushIAy ErzFldSngzDup__WithinT1, ErzFldSngzDup__DupFld_is_fnd(DupFld, Lnxs, T1)
Next
End Function
Private Function ErzFml(Fny$()) As String()
ErzFml = ErzFml__InsideFmlHasInvalidFld(Fny)
End Function
Private Function ErzFml__InsideFmlHasInvalidFld(Fny$()) As String()
Dim J&, Fld$, Fml$, O$(), S$, T1
Dim Lnxs As Lnxs: Lnxs = WLnxszT1("Fml")
For J = 0 To Lnxs.N - 1
    With Lnxs.Ay(J)
        AsgN2tRst .Lin, S, Fld, Fml
        If FstChr(Fml) <> "=" Then
            'PushI O, WMsg_Fml_FstChr(.Lno)
        Else
            Dim ErzFny$(): ErzFny = ErzFnyzFml(Fld, Fml, Fny)
'            PushIAy O, ErzFml__InsideFmlHasInvalidFld1(ErzFny, .Lno, Fld, Fml)
        End If
    End With
Next
ErzFml__InsideFmlHasInvalidFld = O
End Function
Function ErzFnyzFml(Fld$, Fml$, Fny$()) As String() 'Return Subset-Fny (quote by []) in [Fml] which is error. _
It is error if any-FmlFny not in [Fny] or =[Fld]
Dim Ny$(): Ny = NyzMacro(Fml, OpnBkt:="[")
If HasEle(Ny, Fld) Then PushI ErzFnyzFml, Fld
PushIAy ErzFnyzFml, MinusAy(Fml, Fny)
End Function
Private Function ErzFmt() As String()

End Function
Private Function ErzLbl() As String()

End Function
Private Function B_ErzMisFnyzFmti(Fmti) As String()
End Function
Private Function B_ETot_Cnt_Must_1_Fld() As String()

End Function
Private Function B_ETot_Must_Sum_Cnt_Avg(TotT1 As Lnxs) As String()
Dim J%
Dim TotKw$(): TotKw = SyzSS("Avg Cnt Sum")
For J = 0 To TotT1.N - 1
    With TotT1.Ay(J)
'    If Not HasEle(TotKw, .Lin) Then PushI B_ETot_Must_Sum_Cnt_Avg, MTot_Must_Sum_Cnt_Avg
    End With
Next
End Function
Private Function ErzTot_1() '(Cnt$(), CntLxAy%(), Sum$(), SumLxAy%(), Avg$(), AvgLxAy%()) As Erz
'Dim O As New Erz
'Dim J%, C$, Ix%, Msg$
'For J = 0 To UB(Cnt)
'    C = Cnt(J)
'    Ix = Ixy(Sum, C)
'    If Ix >= 0 Then
'        Msg = FmtQQ(M_Fld_IsCnt_FndInSum, CntLxAy(J), Cnt(J), SumLxAy(Ix))
'        O.PushMsg Msg
'    Else
'        Ix = Ixy(Avg, C)
'        If Ix >= 0 Then
'            Msg = FmtQQ(M_Fld_IsCnt_FndInAvg, CntLxAy(J), Cnt(J), AvgLxAy(Ix))
'            O.PushMsg Msg
'        End If
'    End If
'Next
'Set1Lc ErzTot_1 = O
End Function
Private Function B_ELo_() As String()
'W-Erzror-of-LofLinVal:W means working-value. _
which is using the some Module-Lvl-variables and it is private. _
Val here means the LofValFld of LofLin
'E_BLo = Sy(ErzValzNotNum, ErzValzNotInLis, ErzValzFml, ErzValzNotBet)
End Function
Private Function ErzValzFml() As String()

End Function
Private Function B_EVal_NotBet(T1, FmNumVal, ToNumval) As String()
'Dim Lnx(): Lnx = A_T1ToLnxsDic(T1)
End Function
Private Function ErzValzNotInLis() As String()

End Function
Private Function ErzValzNotNum() As String()
Dim T
For Each T In SyzSS("Wdt Lvl")
Next
End Function
Private Function ErzWdt() As String()
End Function
Private Function WLnxszT1(T1) As Lnxs
'If A.KLxs.Exists(T1) Then WLnxszT1 = A.KLxs(T1)
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
'Return Msg if given-Fny has some field not in A.Fny
Dim ErzFny$(): ErzFny = MinusAy(Fny, Fny)
If Si(ErzFny) = 0 Then Exit Function
'PushI WMsgzFny, FmtQQ(M_Fny, ErzFny, Lin_Ty)
End Function
Private Sub Z_ErzBet()
Dim Fny$()
'---------------
Fny = SyzSS("A B")
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
    Act = ErzFldSngzDup(Fny, Lof)
End Sub
Private Sub Z_ErzLof()
Dim Lof$(), LoFny$()
GoSub ZZ
Exit Sub
ZZ:
    Brw ErzLof(Y_Lof, Y_LoFny)
    Return
Tst:
    Act = ErzLof(Lof, LoFny)
    Brw Act
    Stop
    C
    Return
End Sub
Private Function Y_Lof() As String()
Y_Lof = SampLof
End Function
Private Function Y_LoFny() As String()
Y_LoFny = SampLoFny
End Function
Property Get SampLoFny() As String()
SampLoFny = SyzSS("A B C D E F")
End Property
Property Get SampLof() As String()
Erase XX
X "Bet A B C"
X "Lo Nm ABC"
X "Lo Fld A B C D E F G"
X "Ali Left A B"
X "Ali Right D E"
X "Ali Center F"
X "Wdt 10 A B X"
X "Wdt 20 D C C"
X "Wdt 3000 E F G C"
X "Fmt #,## A B C"
X "Fmt #,##.## D E"
X "Lvl 2 A C"
X "Bdr Left A"
X "Bdr Right G"
X "Bdr Col F"
X "Tot Sum A B"
X "Tot Cnt C"
X "Tot Avg D"
X "Tit A abc | sdf"
X "Tit B abc | sdkf | sdfdf"
X "Cor 12345 A B"
X "Fml F A + B"
X "Fml C A * 2"
X "Lbl A lksd flks dfj"
X "Lbl B lsdkf lksdf klsdj f"
SampLof = XX
Erase XX
End Property

Property Get SampLofT1nn$()
SampLofT1nn$ = "Lo Ali Bdr Tot Wdt Fmt Lvl Cor Fml Lbl Tit Bet"
End Property
