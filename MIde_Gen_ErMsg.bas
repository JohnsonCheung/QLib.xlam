Attribute VB_Name = "MIde_Gen_ErMsg"
Option Explicit
Private Type A
    Md As CodeModule
    ErNy() As String
    ErMsgAy() As String
End Type
Private A As A
Const DoczDimItm$ = "DimStmt :: `Dim` DimItm, ..."

Sub GenErMsgMd()
GenErMsgzMd CurMd
End Sub

Private Sub Init(Md As CodeModule)
Dim SrcLy$(): SrcLy = CvSy(AywBetEle(DclLyzMd(Md), "'GenErMsg-Src-Beg.", "'GenErMsg-Src-End."))
Set A.Md = Md
A.ErNy = T1Ay(AyRmvFstChr(SrcLy))
A.ErMsgAy = AyRmvT1(SrcLy)
End Sub

Sub Z_GenErMsgzMd()
Dim Md As CodeModule
GoSub ZZ
Exit Sub
ZZ:
    Const MdVbl$ = "'GenErMsg-Src-Beg." & _
    "|'Val_NotNum      Lno#{Lno&} is [{T1$}] line having Val({Val$}) which should be a number" & _
    "|'Val_NotBet      Lno#{Lno&} is [{T1$}] line having Val({Val$}) which between ({FmNo}) and (ToNm})" & _
    "|'Val_NotInLis    Lno#{Lno&} is [{T1$}] line having invalid Val({ErVal$}).  See valid-value-{VdtValNm$}" & _
    "|'Val_FmlFld      Lno#{Lno&} is [Fml] line having invalid Fml({Fml$}) due to invalid Fny({ErFny$()}).  Valid-Fny are [{VdtFny$()}]" & _
    "|'Val_FmlNotBegEq Lno#{Lno&} is [Fml] line having [{Fml$}] which is not started with [=]" & _
    "|'Fld_NotInFny    Lno#{Lno&} is [{T1$}] line having Fld({F}) which should one of the Fny value.  See [Fny-Value]" & _
    "|'Fld_Dup         Lno#{Lno&} is [{T1$}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno&}" & _
    "|'Fldss_NotSel    Lno#{Lno&} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]" & _
    "|'Fldss_DupSel    Lno#{Lno&} is [{T1$}] line having" & _
    "|'LoNm            Lno#{Lno&} is [Lo-Nm] line having value({Val$}) which is not a good name" & _
    "|'LoNm_Mis        [Lo-Nm] line is missing" & _
    "|'LoNm_Dup        Lno#{Lno&} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno&}" & _
    "|'Tot_DupSel      Lno#{Lno&} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno&} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored." & _
    "|'Bet_3Fld        Lno#{Lno&} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]" & _
    "|'Bet_EqFmTo      Lno#{Lno&} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal." & _
    "|'Bet_FldSeq      Lno#{Lno&} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]" & _
    "|'GenErMsg-Src-End."
    Const MdNm$ = "ZZ_GenMsgzMd"
    RplMd EnsMd(MdNm), RplVbl(MdVbl)
    Set Md = MdzPj(CurPj, MdNm)
    GoSub Tst
    Brw Src(Md)
    RmvMd Md
    Return
Tst:
    GenErMsgzMd Md
    Return
CrtMd:

End Sub

Private Sub GenErMsgzMd(Md As CodeModule)
Init Md
If Sz(A.ErNy) = 0 Then PromptCnl: Exit Sub
RplConstzDic A.Md, ConstDiczErMsg
RplMthzDic A.Md, MthDiczErMsg
End Sub

Sub RplMthzDic(A As CodeModule, MthDic As Dictionary)
Dim MthNm
For Each MthNm In MthDic.Keys
    RplMth A, MthNm, MthDic(MthNm)
Next
End Sub
Sub RplConstzDic(A As CodeModule, ConstDic As Dictionary)

End Sub
Sub RplConst(A As CodeModule, ConstLines)

End Sub
Private Property Get MthDiczErMsg() As Dictionary

End Property

Private Function ConstDiczErMsg() As Dictionary
Const C$ = "Private Const M_?$ = ""?"""
Set ConstDiczErMsg = New Dictionary
Dim ErNm, J%
For Each ErNm In A.ErNy
    ConstDiczErMsg.Add ConstNm(ErNm), FmtQQ(C, ErNm, A.ErMsgAy(J))
    J = J + 1
Next
End Function
Private Function ConstNm$(ErNm)
ConstNm = "M_" & ErNm
End Function
Private Sub Z_MthLineszErMsg()
GoSub T1
Exit Sub
T1:
    A.ErNy = Sy("Val_NotNum")
    A.ErMsgAy = Sy("Lno#{Lno&} is [{T1$}] line having Val({Val$}) which should be a number")
    Ept = ""
    GoTo Tst
Tst:
    Act = MthLineszErMsgs
    D Act
    Stop
    C
    Return
End Sub
Sub Z_Init()
Init Md("MXls_Lof_ErzLof")
D MthLineszErMsgs
Stop
End Sub

Private Property Get MthLineszErMsgs$()
Dim J%, O$()
For J = 0 To UB(A.ErNy)
    PushI O, MthLineszErMsg(A.ErNy(J), A.ErMsgAy(J))
Next
MthLineszErMsgs = JnCrLf(O)
End Property

Private Function MthLineszErMsg$(ErNm$, ErMsg$)
Const C$ = "Private Function Msgz_?$(?)|Msg_? = FmtMacro(""?"", ?)|End Function"
Dim Pm$:           Pm = JnCommaSpc(AywDist(NyzMacro(ErMsg, OpnBkt:="{", InclBkt:=False)))
Dim Calling$: Calling = JnCommaSpc(DimNyzDimItmAy(NyzMacro(ErMsg, OpnBkt:="{", InclBkt:=False)))
MthLineszErMsg = FmtQQ(C, ErNm, Pm, ErNm, ErMsg, Calling)
End Function

Private Sub RmvConstLinzErMsg()
Dim ErNm
For Each ErNm In A.ErNy
    RmvConstzMdPrv A.Md, "M" & ErNm
Next
End Sub

Private Sub RmvMthzErMsg()
Dim ErNm
For Each ErNm In A.ErNy
    RmvMth A.Md, MthNmzErNm(ErNm)
Next
End Sub

Private Function MthNmzErNm$(ErNm)
MthNmzErNm = "Msgz_" & ErNm
End Function

Sub RmvConstzMdPrv(A As CodeModule, ConstNm$)
RmvMdFTIx A, FTIxzMdPrvConst(A, ConstNm)
End Sub

Function FTIxzMdPrvConst(A As CodeModule, ConstNm$) As FTIx
Dim L, Lno&
For Each L In Itr(DclLyzMd(A))
    Lno = Lno + 1
    If HitConstNm(L, ConstNm) Then
        If MthMdy(L) <> "Private" Then Thw CSub, "The given ConstNm should Prv", "Lin ConstNm Lno Md", L, ConstNm, Lno, MdNm(A)
        Set FTIxzMdPrvConst = FTIxzMdLnoCont(A, Lno)
        Exit Function
    End If
Next
Set FTIxzMdPrvConst = EmpFTIx
End Function

