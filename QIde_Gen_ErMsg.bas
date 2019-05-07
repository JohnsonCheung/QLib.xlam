Attribute VB_Name = "QIde_Gen_ErMsg"
Option Explicit
Private Const CMod$ = "MIde_Gen_ErMsg."
Private Const Asm$ = "QIde"
Public Const DoczVdtVerbss$ = "Tag:Definition.  Gen"
Private Type A
    ErNy() As String
    ErMsgAy() As String
End Type
Private A As A
Public Const DoczDimItm$ = "DimStmt :: `Dim` DimItm, ..."
Function X1X$(A As Range)
X1X = TypeName(A.Value)
End Function
Sub GenErMsgzNm(MdNm$)
MdGenErMsg Md(MdNm)
End Sub

Sub GenErMsgMd()
MdGenErMsg CurMd
End Sub

Private Sub Init(Src$())
Dim Dcl$(): Dcl = DclLy(Src)
Dim SrcLy$(): SrcLy = CvSy(AywBetEle(Dcl, "'GenErMsg-Src-Beg.", "'GenErMsg-Src-End."))
'Brw SrcLy, CSub
A.ErNy = T1Sy(RmvFstChrzSy(SrcLy))
A.ErMsgAy = RmvT1zSy(SrcLy)
End Sub

Private Sub Z_SrcGenErMsg()
Dim Src$(), MdNm$
'GoSub T1
'GoSub ZZ1
GoSub ZZ2
Exit Sub
T1:

ZZ2:
    Src = SrczMdNm("MXls_Lof_ErzLof")
    GoSub Tst
    Brw Act
    Return
ZZ1:
    GoSub Set_Src
    MdNm = "XX"
    GoSub Tst
    Brw Act
    Return
Tst:
    Act = SrcGenErMsg(Src, MdNm)
    Return
Set_Src:
    Const X$ = "'GenErMsg-Src-Beg." & _
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
    "|'Bet_N3Fld        Lno#{Lno&} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]" & _
    "|'Bet_EqFmTo      Lno#{Lno&} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal." & _
    "|'Bet_FldSeq      Lno#{Lno&} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]" & _
    "|'GenErMsg-Src-End." & _
    "|Const M_Bet_FldSeq$ = 1"
    Src = SplitVBar(X)
    Return
End Sub

Private Sub A_Prim()
SrcGenErMsg:
End Sub

Private Sub MdGenErMsg(Md As CodeModule)  'eMthNmTy.eeNve
Dim O$(): O = SrcGenErMsg(Src(Md)): 'Brw O: Stop 'Rmk: There is an error when Md is [MXls_Lof_ErzLof].  Er:CannotRmvMth:.
RplMd Md, JnCrLf(O)
End Sub
Private Sub Z_MdGenErMsg()
MdGenErMsg Md("MXls_Lof_ErzLof")
End Sub
Private Sub Z_ErConstDic()
Init ZZ_Src
Brw ErConstDic
End Sub

Function ConstFTIxzMd(A As CodeModule, CnstNm$) As FTIx
Dim L$, I, Lno&
For Each I In Itr(DclLyzMd(A))
    L = I
    Lno = Lno + 1
    If HitCnstNm(L, CnstNm) Then
        If MthMdy(L) <> "Private" Then Thw CSub, "The given CnstNm should Prv", "Lin CnstNm Lno Md", L, CnstNm, Lno, MdNm(A)
        Set ConstFTIxzMd = ContFTIxzMd(A, Lno)
        Exit Function
    End If
Next
Set ConstFTIxzMd = EmpFTIx
End Function

Function ConstFTIx(DclLy$(), CnstNm$) As FTIx
Dim J%
For J = 0 To UB(DclLy)
    If HitCnstNm(DclLy(J), CnstNm) Then
        Set ConstFTIx = ContFTIxzSrc(DclLy, J)
        Exit Function
    End If
Next
Set ConstFTIx = EmpFTIx
End Function

Function SrcGenErMsg(Src$(), Optional MdNm$ = "?") As String()
Init Src
If Si(A.ErNy) = 0 Then Inf CSub, "No GenErMsg-Src-Beg. / GenErMsg-Src-End.", "Md", MdNm: Exit Function
Dim O$(), O1$(), O2$()
O1 = SrcRplConstDic(Src, ErConstDic): 'Brw O1: Stop
O2 = RmvMthInSrc(O1, ErMthNmSet):       'Brw LyzNNAp("MthToRmv BefRmvMth AftRmvMth", ErMthNmSet, O1, O2): Stop
O = AddSy(O2, ErMthLinSy):            'Brw O:Stop
SrcGenErMsg = O
End Function

Function SrcRplConstDic(Src$(), ConstDic As Dictionary) As String()
Dim CnstNm, Dcl$(), Bdy$(), Dcl1$(), Dcl2$()
AsgDclAndBdy Src, Dcl, Bdy
Dcl1 = RmvConstLin(Dcl, KeySet(ConstDic)): 'Brw Dcl1: Stop
'Brw LyzLinesDicItems(ConstDic): Stop
Dcl2 = AddSyAp(Dcl1, LyzLinesDicItems(ConstDic), Bdy): 'Brw Dcl2: Stop
SrcRplConstDic = Dcl2
End Function
Function RmvConstLin(Dcl$(), CnstNmDic As Aset) As String() 'Assume: the const in Dcl to be remove is SngLin
Dim L$, I
For Each I In Itr(Dcl)
    L = I
    If Not HitCnstNmDic(L, CnstNmDic) Then
        PushI RmvConstLin, L
    End If
Next
End Function
Sub AsgDclAndBdy(Src$(), ODcl$(), OBdy$())
Dim J&, F&, U&
U = UB(Src)
F = FstMthIx(Src)
If F < 0 Then
    Erase OBdy
    ODcl = Src
    Exit Sub
End If
For J = 0 To F - 1
    PushI ODcl, Src(J)
Next
For J = F To U
    PushI OBdy, Src(J)
Next
End Sub

Function IupMthByDic(Src$(), ByMthDic As Dictionary) As String() 'Return new Src after replacing Mth in Src by MthDic
IupMthByDic = SplitCrLf(LineszLinesDic(IupDic(MthDic(Src), ByMthDic)))
End Function

Function RplMthzSrc(Src$(), MthNm$, MthLines$) As String()
Dim A() As FTIx: A = MthFTIxAyzSrcMth(Src, MthNm, WiTopRmk:=True)
Dim Ly$(): Ly = SplitCrLf(MthLines)
Select Case Si(A)
Case 0: RplMthzSrc = AddSy(Src, Ly)
Case 1: RplMthzSrc = CvSy(RplAyzMid(Src, A(0), Ly))
Case 2: RplMthzSrc = CvSy(RplAyzMid(Src, A(0), AyeFTIx(Ly, A(1))))
Case Else: Thw CSub, "Error in MthFTIxAyzMth, it should return Sz of 0,1,2, but now it is " & Si(A)
End Select
End Function
Sub IupMthByDicInMd(A As CodeModule, MthDic As Dictionary)
Dim NewSrc$(): NewSrc = IupMthByDic(Src(A), MthDic)
Dim NewLines$: NewLines = JnCrLf(NewSrc)
RplMd A, NewLines
End Sub

Function RplConstInMdByDic(A As CodeModule, ConstDic As Dictionary) As CodeModule
Dim K, CnstNm$
For Each K In ConstDic.Keys
    CnstNm = K
    RplConstInMd A, CnstNm, ConstDic(K)
Next
Set RplConstInMdByDic = A
End Function

Function MdLines(StartLine&, Lines$, Optional InsLno0 = 0) As MdLines
Dim O As New MdLines
Set MdLines = O.Init(StartLine, Lines, InsLno0)
End Function

Function EmpMdLines(A As CodeModule) As MdLines
Dim O As New MdLines
O.InsLno = LnozAftOptzAndImpl(A)
Set EmpMdLines = O
End Function

Function MdLineszMdLno(A As CodeModule, Lno&) As MdLines
Dim Count&, J%
For J = Lno To A.CountOfLines
    Count = Count + 1
    If LasChr(A.Lines(J, 1)) <> "_" Then
        Exit For
    End If
Next
Set MdLineszMdLno = MdLines(Lno, A.Lines(Lno, Count))
End Function

Function MdLineszConst(A As CodeModule, CnstNm$) As MdLines
Dim L$, J&, I
For Each I In DclLy(Src(A))
    L = I
    If HitCnstNm(L, CnstNm) Then
        Set MdLineszConst = MdLineszMdLno(A, J)
        Exit Function
    End If
    J = J + 1
Next
Set MdLineszConst = EmpMdLines(A)
End Function
Sub ThwIfMdLinesMisMatch(A As MdLines, Md As CodeModule)
Dim OldLines$: OldLines = SrcLines(Md)
If OldLines = A.Lines Then
    Thw CSub, "MdLines mis-match with Md", "Md Md-Lines MdLines-Lno MdLines-Lines", MdNm(Md), OldLines, A.StartLine, A.Lines
End If
End Sub
Sub RplLines(A As CodeModule, B As MdLines, NewLines$, Optional LinesNm$ = "MdLines")
Dim OldLinCnt&: OldLinCnt = A.CountOfLines
If B.Count > 0 Then A.DeleteLines B.InsLno, B.Count
A.InsertLines B.InsLno, NewLines
Inf CSub, LinesNm & " is replaced", "Md StartLines NewLinCnt OldLinCnt", MdNm(A), B.StartLine, LinCnt(NewLines), OldLinCnt
End Sub

Sub RplConstInMd(A As CodeModule, CnstNm$, NewLines$)
RplLines A, MdLineszConst(A, CnstNm), NewLines, "MdConst"
End Sub

Private Property Get ErMthNmSet() As Aset
Set ErMthNmSet = AsetzAy(ErMthNy)
End Property

Private Property Get ErMthNy() As String()
Dim ErNy$(): ErNy = A.ErNy
Dim I
For Each I In Itr(ErNy)
    PushI ErMthNy, ErMthNm(I)
Next
End Property

Private Property Get ErConstDic() As Dictionary
Const C$ = "Private Const M_?$ = ""?"""
Set ErConstDic = New Dictionary
Dim ErNm, J%
For Each ErNm In A.ErNy
    ErConstDic.Add ErCnstNm(ErNm), FmtQQ(C, ErNm, A.ErMsgAy(J))
    J = J + 1
Next
End Property

Private Function ErCnstNm$(ErNm)
ErCnstNm = "M_" & ErNm
End Function

Private Sub Z_ErMthLinSy()
'GoSub ZZ
GoSub T1
Exit Sub
ZZ:
    Init ZZ_Src
    Brw ErMthLinSy
    Return
T1:
    A.ErNy = Sy("Val_NotNum")
    A.ErMsgAy = Sy("Lno#{Lno&} is [{T1$}] line having Val({Val$}) which should be a number")
    Ept = Sy("Private Function MsgzVal_NotNum(Lno&, T1$, Val$) As String(): MsgzVal_NotNum = FmtMacro(M_Val_NotNum, Lno, T1, Val): End Function")
    GoTo Tst
Tst:
    Act = ErMthLinSy
    C
    Return
End Sub

Private Sub Z_Init()
Init Src(Md("MXls_Lof_ErzLof"))
If Not HasEle(A.ErNy, "Bet_FldSeq") Then Stop
'Bet_FldSeq
Stop
End Sub

Private Function ZZ_Src() As String()
ZZ_Src = Src(Md("MXls_Lof_ErzLof"))
End Function

Private Function ErMthLinSy() As String() 'One ErMth is one-MulStmtLin
Dim ErNy$(), MsgAy$(), J%, O$()
ErNy = A.ErNy
MsgAy = A.ErMsgAy
For J = 0 To UB(ErNy)
    PushI O, ErMthLinesByNm(ErNy(J), MsgAy(J))
Next
ErMthLinSy = FmtMulStmtSrc(O)
End Function

Private Function ErMthLinesByNm$(ErNm$, ErMsg$)
Dim CNm$:         CNm = ErCnstNm(ErNm)
Dim ErNy$():     ErNy = NyzMacro(ErMsg, ExlBkt:=True)
Dim Pm$:           Pm = JnCommaSpc(AywDist(ErNy))
Dim Calling$: Calling = Jn(AddPfxzSy(DimNyzDimItmAy(ErNy), ", "))
Dim MthNm$:     MthNm = ErMthNm(ErNm)
ErMthLinesByNm = FmtQQ("Private Function ?(?) As String():? = FmtMacro(??):End Function", _
    MthNm, Pm, MthNm, CNm, Calling)
End Function

Private Function ErMthNm$(ErNm)
ErMthNm = "MsgOf_" & ErNm
End Function
