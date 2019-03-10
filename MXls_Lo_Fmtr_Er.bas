Attribute VB_Name = "MXls_Lo_Fmtr_Er"
Option Explicit
Const CMod$ = ""
Const DoczLofValFld As Byte = 1 '

Public Const LofT1nn$ = "Ali Bdr Tot Wdt Fmt Lvl Cor Fml Lbl Tit Bet" ' Fmt. i.tm s.pace s.eparated string
Public Const FmtissSng$ = "                            Fml Lbl Tit Bet" ' Sng.sigle field per line
Const FmtissMul$ = "Ali Bdr Tot Wdt Fmt Lvl Cor                " ' Mul.tiple field per line
Const M01$ = "Lno#? is [?] line having Val(?) which should be a number" 'For Wdt Lvl
Const M02$ = "Lno#? is [?] line having Val(?) which between (?) and (?)" 'For Wdt Lvl
Const M03$ = "Lno#? is [?] line having invalid Val(?).  See valid-value-?"  'For Ali Bdr Tot Cor
Const M04$ = "Lno#? is [Fml] line having invalid Fml(?) due to invalid Fny{?}.  Valid-Fny are [?]." 'For Fml
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
Private A$(), A_Fny$()

Function ErzLof(Lof$(), Fny$()) As String() 'Error-of-ListObj-Formatter:Er.z.Lo.f
Const CSub$ = CMod & "LofEr"
A_Fny = Fny
ErzLof = AyAddAp( _
    WErzVal, WErzFld, WErzFldss, WErzLoNm, _
    WErzAli, WErzBdr, WErzTot, _
    WErzWdt, WErzFmt, WErzLvl, WErzCor, _
    WErzFml, WErzLbl, WErzTit, WErzBet)
End Function

Private Property Get WErzVal() As String() 'W-Error-of-Val:W.Er.z.Val:W means working-value _
|which is using the some Module-Lvl-variables and it is private _
|Val here means the LofValFld of LofLin

WErzVal = SyAddAp(WErzValzNotNum, WErzValzNotInLis, WErzValzFml)
End Property
Private Function WErzValzNotInLis() As String()

End Function
Private Function WErzValzFml() As String()

End Function
Private Function WErzValzNotNum() As String()
End Function

Private Function WErzValzNotBet() As String()
End Function

Private Property Get WErzFld() As String()

End Property

Private Function WErzMisFnyzFmti(Fmti) As String()
'LnxAyzT1 (Fmti)
End Function

'WErzFnyMul W.ork Er.rro z.for Fny. for those fmt line with Mul.tiple fld ------
'DupFny
Private Property Get WErzDupFny() As String()
Dim WErzFny$()
    Dim AlignFny$()
'    AlignFny = AywDist(SSSyzAy(AyRmvTT(Ali)))
'    WErzAli_Fny = AyMinus(AlignFny, A_Fny)
End Property
'Fldss
Private Property Get WErzFldss() As String()

End Property

'LoNm----------------------------------------------------------
Private Property Get WErzLoNm() As String()
WErzLoNm = Sy()
'1Sy(WAli_LeftRightCenter)
End Property

'Ali-----------------------------------------------------------
Private Property Get WErzAli() As String()
WErzAli = Sy(WAli_LeftRightCenter)
End Property

Private Property Get WAli_LeftRightCenter() As String()
'WErzAli_LinEr = WMsgzAliLin(AyeT1Ay(Ali, "Left Right Center"))
End Property

'Bdr-----------------------------------------------------------
Private Function WErzBdr1(X$) As String()
'Return FldAy from Bdr & X
'Dim FldssAy$(): FldssAy = SSSyzAy(AywRmvT1(Bdr, X))
End Function

Private Property Get WErzBdr() As String()
WErzBdr = Sy(WErzBdrExcessFld, WErzBdrExcessLin, WErzBdrDup, WErzBdrFld)
End Property

Private Property Get WErzBdrDup() As String()
'WErzBdrDup = WMsgzDup(AyDupT1(Bdr), Bdr)
End Property

Private Property Get WErzBdrExcessFld() As String()
Dim LFny$(), RfNy$(), CFny$()
LFny = WErzBdr1("Left")
RfNy = WErzBdr1("Right")
CFny = WErzBdr1("Center")
'PushIAy WErzBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, LFny), "Center", "Left")
'PushIAy WErzBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, RfNy), "Center", "Right")
'PushIAy WErzBdrExcessFld, FmtQQ(M_Dup, AyMinus(LFny, RfNy), "Left", "Right")
End Property

Private Property Get WErzBdrExcessLin() As String()
Dim L
'For Each L In Itr(AyeT1Ay(Bdr, "Left Right Center"))
'    PushI WErzBdrExcessLin, FmtQQ(M_Bdr_ExcessLin, L)
'Next
End Property

Private Property Get WErzBdrFld() As String()
Dim Fny$(): Fny = Sy(WErzBdr1("Left"), WErzBdr1("Right"), WErzBdr1("Center"))
WErzBdrFld = WMsgzFny(Fny, "Bdr")
End Property

Private Property Get WErzBet() As String()
WErzBet = Sy(WErzBetDup, WErzBetFny, WErzBetTermCnt)
End Property

Private Property Get WErzBetDup() As String()
'WErzBetDup = WMsgzDup(AyDupT1(Bet), Bet)
End Property

Private Property Get WErzBetFny() As String()
'C$ is the col-c of Bet-line.  It should have 2 item and in Fny
'Return WErz of M_Bet_* if any
End Property

Private Property Get WErzBetTermCnt() As String()
Dim L
'For Each L In Itr(Bet)
    If Sz(SySsl(L)) <> 3 Then
        PushI WErzBetTermCnt, WMsgzBetTermCnt(L, 3)
    End If
'Next
End Property

Private Property Get WErzCor() As String()
Dim L$()
'L = Cor
WErzCor = Sy(WErzCorDup(L), WErzCorFld(L), WErzCorVal(L))
'Cor = L
End Property

Private Function WErzCorDup(IO$()) As String()

End Function

Private Function WErzCorFld(IO$()) As String()

End Function

Private Function WErzCorVal1$(L)
Dim Cor$
Cor = T1(L)
End Function

Private Function WErzCorVal(IO$()) As String()
Dim Msg$(), WErz$(), L
For Each L In IO
    PushI Msg, WErzCorVal1(L)
Next
'IO = AywNoEr(IO, Msg, WErz)
End Function

Private Property Get WErzFml() As String()
WErzFml = Sy(WErzFmlDup, WErzFmlFny)
End Property

Private Property Get WErzFmlDup() As String()
'WErzFmlDup = WMsgzDup(AyDupT1(Fml), Fml)
End Property

Private Property Get WErzFmlFny() As String()
'WErzFmlFny = AyMinus(NyFml(Fml), A_Fny)
End Property

Private Property Get WErzFmt() As String()

End Property

Private Property Get WErzLbl() As String()

End Property

Private Property Get WErzLvl() As String()

End Property

Private Property Get WErzTit() As String()

End Property

Private Property Get WErzTot() As String()
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

Private Property Get WErzTot_1() '(Cnt$(), CntLxAy%(), Sum$(), SumLxAy%(), Avg$(), AvgLxAy%()) As WErz
'Dim O As New WErz
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
'Set1Lc WErzTot_1 = O
End Property

Private Property Get WErzWdt() As String()
End Property

Private Property Get WAny_Tot() As Boolean
Dim Lc As ListColumn
'For Each Lc In A_Lo.ListColumns
    'If LcFmtSpecLy_WAny_Tot(Lc, FmtSpecLy) Then WAny_Tot = True: Exit Function
'Next
End Property

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
Dim WErzFny$(): WErzFny = AyMinus(Fny, A_Fny)
If Sz(WErzFny) = 0 Then Exit Function
'PushI WMsgzFny, FmtQQ(M_Fny, WErzFny, Lin_Ty)
End Function

Private Sub Z_WErzBet()
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
    Act = WErzBet
    C
    Return
End Sub

Function NyzTLinAy(TLinAy$()) As String()
Dim I
For Each I In Itr(TLinAy)
    PushIAy NyzTLinAy, SySsl(I)
Next
End Function

Function FnywLikssAy(Fny$(), LikssAy$()) As String()
Dim F, LikAy$()
LikAy = NyzTLinAy(LikssAy)
For Each F In Itr(Fny)
    If HitLikAy(F, LikAy) Then PushI FnywLikssAy, F
Next
End Function


