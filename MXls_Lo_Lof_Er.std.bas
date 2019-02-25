Attribute VB_Name = "MXls_Lo_Lof_Er"
Option Explicit
Const CMod$ = ""
Private A_Lo As ListObject
Private A_Fny$()
Public Const M_Val_IsNonNum$ = "Lx(?) has Val(?) should be a number"
Public Const M_Val_IsNonLng$ = "Lx(?) has Val(?) should be a 'Long' number"
Public Const M_Val_ShouldBet$ = "Lx(?) has Val(?) should be between [?] and [?]"
Public Const M_Fld_IsInValid$ = "Lx(?) Fld(?) is invalid.  Not found in Fny"
Public Const M_Fld_IsDup$ = "Lx(?) Fld(?) is found duplicated in Lx(?).  This item is ignored"
Public Const M_Nm_LinHasNoVal$ = "Lx(?) is Nm-Lin, it has no value"
Public Const M_Nm_NoNmLin$ = "Nm-Lin is Missing"
Public Const M_Nm_ExcessLin$ = "LX(?) is excess due to Nm-Lin is found above"
Public Const M_Should_Lng$ = "Lx(?) Fld(?) should have val(?) be a long number"
Public Const M_Should_Num$ = "Lx(?) Fld(?) should have val(?) be a number"
Public Const M_Should_Bet$ = "Lx(?) Fld(?) should have val(?) be between (?) and (?)"

Const M_Fny$ = "Lin_Ty(?) has these Fld(?) in not Fny"
Const M_Bdr_ExcessFld$ = "These Fld(?) in [Bdr ?] already Has in [Bdr ?], they are skipped in setting border"
Const M_Bdr_ExcessLin$ = "These Fld(?) in [Bdr ?] already Has in [Bdr ?], they are skipped in setting border"
Const M_CorVal$ = "In Lin(?)-Color(?), color cannot convert to long"
Const M_Fld_IsAvg_FndInSum$ = "Lin(?)-Fld(?), which is TAvg-Fld, but also found in TSum-Lx(?)"
Const M_Fld_IsCnt_FndInSum$ = "Lin(?)-Fld(?), which is TCnt-Fld, but also found in TSum-Lx(?)"
Const M_Fld_IsCnt_FndInAvg$ = "Lin(?)-Fld(?), which is TCnt-Fld, but also found in TAvg-Lx(?)"
Const M_Bet_Should2Term = "Lin(?)-Fld(?) is Bet-Line.  It should have 2 terms"
Const M_Bet_InvalidTerm = "Lin(?)-Fld(?) is Bet-Line.  It has invalid term(?)"
Const M_Dup$ = "Lin(?)-Fld(?) is duplicated.  The line is skipped"
Private A As Dictionary

Private Property Get WErzAli() As String()
WErzAli = Sy(WErzAlignLin, WErzAlignFny)
End Property

Private Property Get WErzAlignFny() As String()
Dim WErzFny$()
    Dim AlignFny$()
'    AlignFny = AywDist(SSSyzAy(AyRmvTT(Ali)))
'    WErzFny = AyMinus(AlignFny, A_Fny)
End Property

Private Property Get WErzAlignLin() As String()
'WErzAlignLin = WMsgzAliLin(AyeT1Ay(Ali, "Left Right Center"))
End Property

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
PushIAy WErzBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, LFny), "Center", "Left")
PushIAy WErzBdrExcessFld, FmtQQ(M_Dup, AyMinus(CFny, RfNy), "Center", "Right")
PushIAy WErzBdrExcessFld, FmtQQ(M_Dup, AyMinus(LFny, RfNy), "Left", "Right")
End Property

Private Property Get WErzBdrExcessLin() As String()
Dim L
'For Each L In Itr(AyeT1Ay(Bdr, "Left Right Center"))
    PushI WErzBdrExcessLin, FmtQQ(M_Bdr_ExcessLin, L)
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
Function LofEr(LoFmtrDic As Dictionary) As String()
Const CSub$ = CMod & "LofEr"
Set A = LoFmtrDic
LofEr = AyAddAp( _
    WErzAli, WErzBdr, WErzTot, _
    WErzWdt, WErzFmt, WErzLvl, WErzCor, _
    WErzFml, WErzLbl, WErzTit, WErzBet)
End Function


Private Function WMsgzAliLin(Ly$()) As String()
If Sz(Ly) Then Exit Function
End Function

Private Function WMsgzBetTermCnt$(L, NTerm%)

End Function

Private Function WMsgzDup1(N, Ly$()) As String()
Dim L
For Each L In Ly
    If T1(L) = N Then PushI WMsgzDup1, FmtQQ(M_Dup, L, N)
Next
End Function

Private Function WMsgzDup(DupNy$(), Ly$()) As String()
Dim N
For Each N In Itr(DupNy)
    PushIAy WMsgzDup, WMsgzDup1(N, Ly)
Next
End Function

Private Function WMsgzFny(Fny$(), Lin_Ty$) As String()
'Return Msg if given-Fny has some field not in A_Fny
Dim WErzFny$(): WErzFny = AyMinus(Fny, A_Fny)
If Sz(WErzFny) = 0 Then Exit Function
PushI WMsgzFny, FmtQQ(M_Fny, WErzFny, Lin_Ty)
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


