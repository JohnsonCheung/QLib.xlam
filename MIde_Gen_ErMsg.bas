Attribute VB_Name = "MIde_Gen_ErMsg"
Option Explicit
Private Type A
    ErNy() As String
    ErMsgAy() As String
End Type
Private A As A
Const ™DimItm$ = "DimStmt :: `Dim` DimItm, ..."
Sub GenErMsg(MdNm$)
GenErMsgzMd Md(MdNm)
End Sub

Sub GenErMsgMd()
GenErMsgzMd CurMd
End Sub

Private Sub Init(Md As CodeModule)
Dim Dcl$(): Dcl = DclLyzMd(Md)
'Brw Dcl
Dim SrcLy$(): SrcLy = CvSy(AywBetEle(Dcl, "'GenErMsg-Src-Beg.", "'GenErMsg-Src-End."))
'Brw SrcLy, CSub
A.ErNy = T1Ay(AyRmvFstChr(SrcLy))
A.ErMsgAy = AyRmvT1(SrcLy)
End Sub

Sub Z_GenErMsgzMd()
Const MdNm$ = "ZZ_GenMsgzMd"
Dim Md As CodeModule
GoSub T1
'GoSub ZZ
Exit Sub
T1:
    Set Md = MdzPj(CurPj, "MXls_Lof_ErzLof")
    GoTo Tst
ZZ:
    GoSub CrtMd
    Set Md = MdzPj(CurPj, MdNm)
    GoSub Tst
    Brw Src(Md)
    RmvMd Md
    Return
Tst:
    GenErMsgzMd Md
    Return
CrtMd:
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
    RplMd EnsMd(MdNm), RplVbl(MdVbl)
    Return
End Sub

Private Sub GenErMsgzMd(Md As CodeModule)
Init Md
If Si(A.ErNy) = 0 Then Inf CSub, "No GenErMsg-Src-Beg. / GenErMsg-Src-End.", "Md", MdNm(Md): Exit Sub
Brw SrcRplConstzDic(Src(Md), ErMsgConstDic)
Stop
Dim O$(): O = SrcRplMthzDic(SrcRplConstzDic(Src(Md), ErMsgConstDic), MthDiczErMsg)
Stop
Brw O
Exit Sub
RplMd Md, JnCrLf(O)
End Sub

Sub Z_ErMsgConstDic()
Init Md("MXls_Lof_ErzLof")
Brw ErMsgConstDic
End Sub

Function ConstFTIxzMd(A As CodeModule, ConstNm$) As FTIx
Dim L, Lno&
For Each L In Itr(DclLyzMd(A))
    Lno = Lno + 1
    If HitConstNm(L, ConstNm) Then
        If MthMdy(L) <> "Private" Then Thw CSub, "The given ConstNm should Prv", "Lin ConstNm Lno Md", L, ConstNm, Lno, MdNm(A)
        Set ConstFTIxzMd = ContFTIxzMd(A, Lno)
        Exit Function
    End If
Next
Set ConstFTIxzMd = EmpFTIx
End Function

Function ConstFTIx(DclLy$(), ConstNm) As FTIx
Dim J%
For J = 0 To UB(DclLy)
    If HitConstNm(DclLy(J), ConstNm) Then
        Set ConstFTIx = ContFTIxzSrc(DclLy, J)
        Exit Function
    End If
Next
Set ConstFTIx = EmpFTIx
End Function

Function SrcRplConstzDic(Src$(), ConstDic As Dictionary) As String()
Dim ConstNm, Dcl$(), Bdy$()
AsgSrcToDclAndBdy Src, Dcl, Bdy
For Each ConstNm In ConstDic.Keys
    Dcl = DclRplConst(Dcl, ConstNm, ConstDic(ConstNm))
Next
SrcRplConstzDic = SyAdd(Dcl, Bdy)
End Function

Sub AsgSrcToDclAndBdy(Src$(), ODcl$(), OBdy$())
Dim J&, F&
F = FstMthIx(Src)
For J = 0 To F - 1
    PushI ODcl, Src(J)
Next
For J = F To U
    PushI OBdy, Src(J)
Next
End Sub

Function DclRplConst(Dcl$(), ConstNm, ConstLines$) As String()
SrcRplConst = CvSy(AyAdd(AyeFTIx(Src, ConstFTIx(Src, ConstNm)), SplitCrLf(ConstLines)))
End Function

Function SrcRplMthzDic(Src$(), MthDic As Dictionary) As String()
SrcRplMthzDic = SyAdd(SrcExlMth(Src, KeySet(MthDic)), LyzLinesDicByItems(MthDic))
End Function

Function SrcRplMth(Src$(), MthNm, MthLines) As String()

End Function

Sub RplMthzDic(A As CodeModule, MthDic As Dictionary)
Dim MthNm
For Each MthNm In MthDic.Keys
    RplMth A, MthNm, MthDic(MthNm)
Next
End Sub

Sub RplConstzDic(A As CodeModule, ConstDic As Dictionary)
Dim K
For Each K In ConstDic.Keys
    RplConst A, K, ConstDic(K)
Next
End Sub

Function MdLines(StartLine, Lines, Optional InsLno0 = 0) As MdLines
Dim O As New MdLines
Set MdLines = O.Init(StartLine, Lines, InsLno0)
End Function

Function EmpMdLines(A As CodeModule) As MdLines
Dim O As New MdLines
O.InsLno = LnozAftOpt(A)
Set EmpMdLines = O
End Function

Function MdLineszMdLno(A As CodeModule, Lno) As MdLines
Dim Count&, J%
For J = Lno To A.CountOfLines
    Count = Count + 1
    If LasChr(A.Lines(J, 1)) <> "_" Then
        Exit For
    End If
Next
Set MdLineszMdLno = MdLines(Lno, A.Lines(Lno, Count))
End Function

Function MdLineszConst(A As CodeModule, ConstNm) As MdLines
Dim L, J%
For Each L In DclLy(Src(A))
    If HitConstNm(L, ConstNm) Then
        Set MdLineszConst = MdLineszMdLno(A, J)
        Exit Function
    End If
    J = J + 1
Next
Set MdLineszConst = EmpMdLines(A)
End Function

Sub RplMdLines(A As CodeModule, B As MdLines, NewLines, Optional LinesNm$ = "MdLines")
Dim OldLines$: If B.Count > 0 Then OldLines = A.Lines(B.StartLine, B.Count)
If OldLines = NewLines Then Inf CSub, "Same " & LinesNm, "Md StartLine Count FstLin", MdNm(A), B.StartLine, B.Count, FstLin(B.Lines): Exit Sub
If B.Count > 0 Then A.DeleteLines B.InsLno, B.Count
A.InsertLines B.InsLno, NewLines
Inf CSub, LinesNm & " is replaced", "Md StartLines NewLinCnt OldLinCnt NewLines OldLines", MdNm(A), B.StartLine, LinCnt(NewLines), B.Count, NewLines, OldLines
End Sub

Sub RplConst(A As CodeModule, ConstNm, NewLines)
RplMdLines A, MdLineszConst(A, ConstNm), NewLines, "MdConst"
End Sub

Private Property Get MthDiczErMsg() As Dictionary 'Key is ErMsgMthNm and Val is ErMsgMthLines
Dim N, MthNm$, MthLines$, J%
Set MthDiczErMsg = New Dictionary
For Each N In Itr(A.ErNy)
    MthNm = ErMsgMthNm(N)
    MthLines = ErMsgMthLines(CStr(N), A.ErMsgAy(J))
    MthDiczErMsg.Add MthNm, MthLines
    J = J + 1
Next
End Property

Private Property Get ErMsgConstDic() As Dictionary
Const C$ = "Private Const M_?$ = ""?"""
Set ErMsgConstDic = New Dictionary
Dim ErNm, J%
For Each ErNm In A.ErNy
    ErMsgConstDic.Add ConstNm(ErNm), FmtQQ(C, ErNm, A.ErMsgAy(J))
    J = J + 1
Next
End Property

Private Function ConstNm$(ErNm)
ConstNm = "M_" & ErNm
End Function

Private Sub Z_ErMsgMthDic()
GoSub T1
Exit Sub
T1:
    A.ErNy = Sy("Val_NotNum")
    A.ErMsgAy = Sy("Lno#{Lno&} is [{T1$}] line having Val({Val$}) which should be a number")
    Ept = ""
    GoTo Tst
Tst:
    Set Act = ErMsgMthDic
    Brw Act
    Stop
    C
    Return
End Sub

Sub Z_Init()
Init Md("MXls_Lof_ErzLof")
D ErMsgMthDic
Stop
End Sub

Private Property Get ErMsgMthDic() As Dictionary
Dim J%, ErNy$()
ErNy = A.ErNy
Set ErMsgMthDic = New Dictionary
For J = 0 To UB(ErNy)
    ErMsgMthDic.Add ErNy(J), ErMsgMthLines(ErNy(J), A.ErMsgAy(J))
Next
End Property

Private Function ErMsgMthLines$(ErNm$, ErMsg$)
Const C$ = "Private Function Msgz_?(?) As String()|Msgz_? = FmtMacro(""?"", ?)|End Function"
Dim ErNy$():     ErNy = NyzMacro(ErMsg, ExlBkt:=True)
Dim Pm$:           Pm = JnCommaSpc(AywDist(ErNy))
Dim Calling$: Calling = JnCommaSpc(DimNyzDimItmAy(ErNy))
ErMsgMthLines = FmtQQ(C, ErNm, Pm, ErNm, ErMsg, Calling)
End Function

Private Function ErMsgMthNm$(ErNm)
ErMsgMthNm = "ErMsgz_" & ErNm
End Function

Sub RmvConst(A As CodeModule, ConstNm$)
RmvMdFTIx A, ConstFTIxzMd(A, ConstNm)
End Sub

