Attribute VB_Name = "QXls_Fun_ErzLof"
Option Compare Text
Option Explicit
Private Const Asm$ = "QXls"
Private Const CMod$ = "MXls_Lof_ErzLof."
Public Const LofT1nn$ = _
                            "Lo Ali Bdr Tot Wdt Fmt Lvl Cor Fml Lbl Tit Bet" ' Fmt. i.tm s.pace s.eparated string
Public Const LofT1nnzSng$ = "                               Fml Lbl Tit Bet" ' Sng.sigle field per line
Public Const LofT1nnzMul$ = "Lo Ali Bdr Tot Wdt Fmt Lvl Cor                " ' Mul.tiple field per line
'
'Lo  Nm  Er    [Lo Nm] has error
'Lo  Nm  Mis   [Lo Nm] line is missed
'Lo  Nm  Dup   [Lo Nm] is Dup
'Lo  Fny Mis   [Lo Fny] is missed
'Lo  Fny Dup   [Lo Fny] is missed
'Ali Val NLis  [Ali Val] is not in @AliVal
'Ali Fld NLis  [Ali Fld] is not in @LoFny
'Bdr Val NLis  [Bdr Val] is not in @BdrVal
'Tot Val NLis  [Tot Val] is not in @TotVal
'Wdt Val NNum  [Wdt Val] is not number
'Wdt Val Mis   [Wdt Val] is missed
'Wdt Val NBet  [Wdt Val] is not between 3 to 100
'Lvl Val NNum  [Lvl Val] is not a number
'Lvl Val NBet  [Lvl Val] is not between 2 and 8
'Lvl Fld NLis  [Lvl Fld] is not in @LoFny
'Lvl Fld Dup
'

Function LofT1Ny() As String()
LofT1Ny = TermAy(LofT1nn)
End Function

Private Function XELoFldDup(DoLo As Drs) As String()

End Function

Private Function XLoNmMis(LoLy$()) As Boolean

End Function
Private Function XELoNmMis(DoLo As Drs) As String()
'If IsLoMisNm Then XELoNmMis = Sy("No LoNm")
End Function

Private Function XELoNmEr(DoLo As Drs) As String()
'XELoNmEr = M_Lo_ErNm(LnoAy)
End Function

Private Function XELoNmDup(DoLo As Drs) As String()
'Dim Lnoss: For Each Lnoss In Itr(LnossAy)
'    PushI XELoNmDup, FmtQQ(C_Lo_ErNm, Lnoss)
'Next
End Function

Private Function XELoFldMis(DoLo As Drs) As String()

End Function
Private Function XDoLo(LTD As Drs) As Drs

End Function
Private Sub Z_ErzLof()
Dim Lof$(), LofNy$()
GoSub T0
T0:
    Lof = SampLof
    LofNy = SyzSS("A B C D E F G")
    Ept = Sy()
    GoTo Tst
Tst:
    Act = ErzLof(Lof, LofNy)
    C
    Return
End Sub
Function ErzLof(Lof$(), LofNy$()) As String()
':Lof: :Fmtr #ListObj-Fmtr# !
':Fmtr: :Ly #Formatter#
Dim A$(), B$(), C$(), D$(), E$()
Dim LTD As Drs: LTD = DoLTD(Lof)
Dim DoLo As Drs: DoLo = XDoLo(LTD)
                A = XELoNmMis(DoLo)
                B = XELoNmEr(DoLo)
                C = XELoNmDup(DoLo)
                D = XELoFldMis(DoLo)
                E = XELoFldDup(DoLo)
Dim ELo$():   ELo = Sy(A, B, C, D, E)

'                A = Sy(FmtQQ(MAli_MustLRCenter))
Dim EAli$(): EAli = Sy(A)

                A = XEVal_NotBet("Wdt", 10, 200)
Dim EWdt$(): EWdt = Sy(A)


Dim EFmt$(): EFmt = Sy(A, B, C)

                B = XEVal_NotBet("Lvl", 2, 9)
Dim ELvl$(): ELvl = Sy(A, B, C)


Dim ECor$(): ECor = Sy(A, B, C)

Dim EFml$(): EFml = Sy(A, B, C)


Dim ELbl$(): ELbl = Sy(A, B, C)

Dim ETit$(): ETit = Sy(A, B, C)

Dim EBet$(): EBet = Sy(A)

Dim EBdr$(): EBdr = Sy(A, B)

Dim ETot$(): ETot = Sy(A, B, C)

ErzLof = Sy(ELo, EAli, EBdr, ETot, EWdt, EFmt, ELvl, ECor, EFml, ELbl, ETit, EBet)
End Function

Private Function WAny_Tot() As Boolean
Dim LC As ListColumn
'For Each Lc In A_Lo.ListColumns
    'If LcFmtSpecLy_WAny_Tot(Lc, FmtSpecLy) Then WAny_Tot = True: Exit Function
'Next
End Function
Private Function ErzBdr1(X$) As String()
'Return FldAy from Bdr & X
'Dim FldssAy$(): FldssAy = SSSyzAy(AwRmvT1(Bdr, X))
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
'IO = AwNoErz(IO, Msg, Erz)
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

Private Function ErzFldSngzDup__WithinT1(T1) As String() 'Within T1 any fld is dup?
Dim DupFld$, I

'For Each I In Itr(DupT2AyzLnxs())
    DupFld = I
'    PushIAy ErzFldSngzDup__WithinT1, ErzFldSngzDup__DupFld_is_fnd(DupFld, Lnxs, T1)
'Next
End Function

Private Function XErFml(Fny$()) As String()
XErFml = XErFml__InsideFmlHasInvalidFld(Fny)
End Function

Private Function XErFml__InsideFmlHasInvalidFld(Fny$()) As String()
Dim J&, Fld$, Fml$, O$(), S$, T1
'Dim Lnxs As Lnxs: Lnxs = WLnxszT1("Fml")
'For J = 0 To Lnxs.N - 1
    'With Lnxs.Ay(J)
'        AsgTTRst .Lin, S, Fld, Fml
        If FstChr(Fml) <> "=" Then
            'PushI O, WMsg_Fml_FstChr(.Lno)
        Else
            Dim ErzFny$(): 'ErzFny = ErzFmlFld(Fld, Fml, Fny)
'            PushIAy O, ErzFml__InsideFmlHasInvalidFld1(ErzFny, .Lno, Fld, Fml)
        End If
    'End With
'Next
XErFml__InsideFmlHasInvalidFld = O
End Function

Private Function ErzFmlFld(Fld$, Fml$, Fny$()) As String()
'Ret :urn Subset-Fny (quote by []) in [Fml] which is error. _
It is error if any-FmlFny not in [Fny] or =[Fld]
Dim Ny$(): Ny = NyzMacro(Fml, OpnBkt:="[")
If HasEle(Ny, Fld) Then 'PushI ErzFmlFld, Fld
'PushIAy ErzFmlFld, MinusAy(Fml, Fny)
End If
End Function

Private Function ErzFmt() As String()

End Function
Private Function ErzLbl() As String()

End Function
Private Function B_ErzMisFnyzFmti(Fmti) As String()
End Function

Private Function B_ETot_Cnt_Must_1_Fld() As String()

End Function

Private Function B_ETot_Must_Sum_Cnt_Avg() As String()
Dim J%
Dim TotKw$(): TotKw = SyzSS("Avg Cnt Sum")
'For J = 0 To TotT1.N - 1
'    With TotT1.Ay(J)
'    If Not HasEle(TotKw, .Lin) Then PushI B_ETot_Must_Sum_Cnt_Avg, MTot_Must_Sum_Cnt_Avg
'    End With
'Next
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
'E_BLo = Sy(ErzVzNotNum, ErzVzNotInLis, ErzVzFml, ErzVzNotBet)
End Function
Private Function ErzVzFml() As String()

End Function
Private Function XEVal_NotBet(T1, FmNumVal, ToNumval) As String()
'Dim Lnx(): Lnx = A_T1ToLnxsDic(T1)
End Function
Private Function ErzVzNotInLis() As String()

End Function
Private Function ErzVzNotNum() As String()
Dim T
For Each T In SyzSS("Wdt Lvl")
Next
End Function
Private Function ErzWdt() As String()
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

Private Function Y_Lof() As String()
Y_Lof = SampLof
End Function

Private Function Y_LoFny() As String()
Y_LoFny = FoSampLo
End Function

Property Get FoSampLo() As String()
FoSampLo = SyzSS("A B C D E F")
End Property

Property Get SampLof() As String()
Dim A As New Bfr
With A
.Var "Sum Bet  A B C"
.Var "Lo Nm BC"
.Var "Lo Fld  B C D E F G"
.Var "Ali Left  B"
.Var "Ali Right D E"
.Var "Ali Center F"
.Var "Wdt 10  B X"
.Var "Wdt 20 D C C"
.Var "Wdt 3000 E F G C"
.Var "Fmt #,##  B C"
.Var "Fmt #,##.## D E"
.Var "Lvl 2  C"
.Var "Bdr Left "
.Var "Bdr Right G"
.Var "Bdr Center F"
.Var "Tot Sum  B"
.Var "Tot Cnt C"
.Var "Tot Avg D"
.Var "Tit A bc | sdf"
.Var "Tit B bc | sdkf | sdfdf"
.Var "Cor 12345  B"
.Var "Fml F  A + B"
.Var "Fml C  B * 2"
.Var "Lbl A lksd flks dfj"
.Var "Lbl B lsdkf lksdf klsdj f"
End With
SampLof = AlignLyzTTRst(A.Ly)
End Property

Function Lnoss$(Ixy() As Long)
Lnoss = JnSpc(AyIncEle1(Ixy))
End Function

Private Sub Z()
QXls_Fun_ErzLof:
End Sub
