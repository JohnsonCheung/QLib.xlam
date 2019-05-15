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
Sub GenErMsgzNm(Mdn)
MdGenErMsg Md(Mdn)
End Sub

Sub GenErMsgMd()
MdGenErMsg CMd
End Sub

Private Sub Init(Src$())
Dim Dcl$(): Dcl = DclLy(Src)
Dim SrcLy$(): SrcLy = CvSy(AywBetEle(Dcl, "'GenErMsg-Src-Beg.", "'GenErMsg-Src-End."))
'Brw SrcLy, CSub
A.ErNy = T1Ay(RmvFstChrzSy(SrcLy))
A.ErMsgAy = RmvT1zAy(SrcLy)
End Sub

Private Sub Z_SrcGenErMsg()
Dim Src$(), Mdn
'GoSub T1
'GoSub ZZ1
GoSub ZZ2
Exit Sub
T1:

ZZ2:
    Src = SrczMdn("MXls_Lof_ErzLof")
    GoSub Tst
    Brw Act
    Return
ZZ1:
    GoSub Set_Src
    Mdn = "XX"
    GoSub Tst
    Brw Act
    Return
Tst:
    Act = SrcGenErMsg(Src, Mdn)
    Return
Set_Src:
    Const X$ = "'GenErMsg-Src-Beg." & _
    "|'Val_NotNum      Lno#{Lno} is [{T1$}] line having Val({Val$}) which should be a number" & _
    "|'Val_NotBet      Lno#{Lno} is [{T1$}] line having Val({Val$}) which between ({FmNo}) and (ToNm})" & _
    "|'Val_NotInLis    Lno#{Lno} is [{T1$}] line having invalid Val({ErVal$}).  See valid-value-{VdtValNm$}" & _
    "|'Val_FmlFld      Lno#{Lno} is [Fml] line having invalid Fml({Fml$}) due to invalid Fny({ErFny$()}).  Valid-Fny are [{VdtFny$()}]" & _
    "|'Val_FmlNotBegEq Lno#{Lno} is [Fml] line having [{Fml$}] which is not started with [=]" & _
    "|'Fld_NotInFny    Lno#{Lno} is [{T1$}] line having Fld({F}) which should one of the Fny value.  See [Fny-Value]" & _
    "|'Fld_Dup         Lno#{Lno} is [{T1$}] line having Fld({F}) which is duplicated and ignored due to it has defined in Lno#{AlreadyInLno}" & _
    "|'Fldss_NotSel    Lno#{Lno} is [{T1$}] line having Fldss({Fldss$}) which should select one for Fny value.  See [Fny-Value]" & _
    "|'Fldss_DupSel    Lno#{Lno} is [{T1$}] line having" & _
    "|'LoNm            Lno#{Lno} is [Lo-Nm] line having value({Val$}) which is not a good name" & _
    "|'LoNm_Mis        [Lo-Nm] line is missing" & _
    "|'LoNm_Dup        Lno#{Lno} is [Lo-Nm] which is duplicated and ignored due to there is already a [Lo-Nm] in Lno#{AlreadyInLno}" & _
    "|'Tot_DupSel      Lno#{Lno} is [Tot-{TotKd$}] line having Fldss({Fldss$}) selecting SelFld({SelFld$}) which is already selected by Lno#{AlreadyInLno} of [Tot-{AlreadyTotKd$}].  The SelFld is ignored." & _
    "|'Bet_N3Fld        Lno#{Lno} is [Bet] line.  It should have 3 fields, but now it has (?) fields of [?]" & _
    "|'Bet_EqFmTo      Lno#{Lno} is [Bet] line and ignored due to FmFld(?) and ToFld(?) are equal." & _
    "|'Bet_FldSeq      Lno#{Lno} is [Bet] line and ignored due to Fld(?), FmFld(?) and ToFld(?) are not in order.  See order the Fld, FmFld and ToFld in [Fny-Value]" & _
    "|'GenErMsg-Src-End." & _
    "|Const M_Bet_FldSeq$ = 1"
    Src = SplitVBar(X)
    Return
End Sub

Private Sub A_Prim()
SrcGenErMsg:
End Sub

Private Sub MdGenErMsg(Md As CodeModule)  'eMthnTy.eeNve
Dim O$(): O = SrcGenErMsg(Src(Md)): 'Brw O: Stop 'Rmk: There is an error when Md is [MXls_Lof_ErzLof].  Er:CannotRmvMth:.
RplMd Md, JnCrLf(O)
End Sub
Private Sub Z_MdGenErMsg()
MdGenErMsg Md("MXls_Lof_ErzLof")
End Sub
Private Sub Z_ErConstDic()
'Init Y_Src
Brw ErConstDic
End Sub

Function ConstFeizMd(A As CodeModule, Cnstn$) As Fei
Dim L$, I, Lno
For Each I In Itr(DclLyzM(A))
    L = I
    Lno = Lno + 1
    If HitCnstn(L, Cnstn) Then
        If MthMdy(L) <> "Private" Then Thw CSub, "The given Cnstn should Prv", "Lin Cnstn Lno Md", L, Cnstn, Lno, Mdn(A)
        'ConstFeizMd = ContFeizMd(A, Lno)
        Exit Function
    End If
Next
End Function

Function ConstFei(DclLy$(), Cnstn$) As Fei
Dim J&
For J = 0 To UB(DclLy)
    If HitCnstn(DclLy(J), Cnstn) Then
        'ConstFei = ContFeizS(DclLy, J)
        Exit Function
    End If
Next
End Function

Function SrcGenErMsg(Src$(), Optional Mdn = "?") As String()
Init Src
If Si(A.ErNy) = 0 Then Inf CSub, "No GenErMsg-Src-Beg. / GenErMsg-Src-End.", "Md", Mdn: Exit Function
Dim O$(), O1$(), O2$()
O1 = SrcRplConstDic(Src, ErConstDic): 'Brw O1: Stop
O2 = RmvMthInSrc(O1, ErMthnSet):       'Brw LyzNNAp("MthToRmv BefRmvMth AftRmvMth", ErMthnSet, O1, O2): Stop
O = AddSy(O2, ErMthLiny):            'Brw O:Stop
SrcGenErMsg = O
End Function

Function SrcRplConstDic(Src$(), ConstDic As Dictionary) As String()
Dim Cnstn, Dcl$(), Bdy$(), Dcl1$(), Dcl2$()
AsgDclAndBdy Src, Dcl, Bdy
Dcl1 = RmvConstLin(Dcl, KeySet(ConstDic)): 'Brw Dcl1: Stop
'Brw LyzLinesDicItems(ConstDic): Stop
Dcl2 = Sy(Dcl1, LyzLinesDicItems(ConstDic), Bdy): 'Brw Dcl2: Stop
SrcRplConstDic = Dcl2
End Function
Function RmvConstLin(Dcl$(), CnstnDic As Aset) As String() 'Assume: the const in Dcl to be remove is SngLin
Dim L$, I
For Each I In Itr(Dcl)
    L = I
    If Not HitCnstnDic(L, CnstnDic) Then
        PushI RmvConstLin, L
    End If
Next
End Function
Sub AsgDclAndBdy(Src$(), ODcl$(), OBdy$())
Dim J&, F&, U&
U = UB(Src)
F = FstMthIxzS(Src)
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
Function DltMthzSN(Src$(), Mthn) As String()

End Function

Function RplMthzSNL(Src$(), Mthn, MthLines$) As String()
Dim O$()
Dim Ix&: 'Ix = MthIxzSrcNmTy SN(Src,Mthn)
O = DltMthzSN(Src, Mthn)
'O = InsMthzSIL(Src, Ix, MthLines)
RplMthzSNL = O
End Function
Sub IupMthByDicM(A As CodeModule, MthDic As Dictionary)
Dim NewSrc$(): NewSrc = IupMthByDic(Src(A), MthDic)
Dim NewLines$: NewLines = JnCrLf(NewSrc)
RplMd A, NewLines
End Sub

Function RplConstMByDic(A As CodeModule, ConstDic As Dictionary) As CodeModule
Dim K, Cnstn$
For Each K In ConstDic.Keys
    Cnstn = K
    RplConstM A, Cnstn, ConstDic(K)
Next
Set RplConstMByDic = A
End Function

Sub RplConstM(A As CodeModule, Cnstn$, NewLines$)
RplLines A, MdLineszConst(A, Cnstn), NewLines, "MdConst"
End Sub

Private Property Get ErMthnSet() As Aset
Set ErMthnSet = AsetzAy(ErMthny)
End Property

Private Property Get ErMthny() As String()
Dim ErNy$(): ErNy = A.ErNy
Dim I
For Each I In Itr(ErNy)
    PushI ErMthny, ErMthn(I)
Next
End Property

Private Property Get ErConstDic() As Dictionary
Const C$ = "Private Const M_?$ = ""?"""
Set ErConstDic = New Dictionary
Dim ErNm, J%
For Each ErNm In A.ErNy
    ErConstDic.Add ErCnstn(ErNm), FmtQQ(C, ErNm, A.ErMsgAy(J))
    J = J + 1
Next
End Property

Private Function ErCnstn$(ErNm)
ErCnstn = "M_" & ErNm
End Function

Private Sub Z_ErMthLiny()
'GoSub ZZ
GoSub T1
Exit Sub
ZZ:
    Init Y_Src
    Brw ErMthLiny
    Return
T1:
    A.ErNy = Sy("Val_NotNum")
    A.ErMsgAy = Sy("Lno#{Lno} is [{T1$}] line having Val({Val$}) which should be a number")
    Ept = Sy("Private Function MsgzVal_NotNum(Lno, T1, Val$) As String(): MsgzVal_NotNum = FmtMacro(M_Val_NotNum, Lno, T1, Val): End Function")
    GoTo Tst
Tst:
    Act = ErMthLiny
    C
    Return
End Sub

Private Sub Z_Init()
Init Src(Md("MXls_Lof_ErzLof"))
If Not HasEle(A.ErNy, "Bet_FldSeq") Then Stop
'Bet_FldSeq
Stop
End Sub

Private Function Y_Src() As String()
Y_Src = Src(Md("MXls_Lof_ErzLof"))
End Function

Private Function ErMthLiny() As String() 'One ErMth is one-MulStmtLin
Dim ErNy$(), MsgAy$(), J%, O$()
ErNy = A.ErNy
MsgAy = A.ErMsgAy
For J = 0 To UB(ErNy)
    PushI O, ErMthLinesByNm(ErNy(J), MsgAy(J))
Next
ErMthLiny = FmtMulStmtSrc(O)
End Function

Private Function ErMthLinesByNm$(ErNm$, ErMsg$)
Dim CNm$:         CNm = ErCnstn(ErNm)
Dim ErNy$():     ErNy = NyzMacro(ErMsg)
Dim Pm$:           Pm = JnCommaSpc(AywDist(ErNy))
Dim Calling$: Calling = Jn(AddPfxzAy(DimNyzDimItmAy(ErNy), ", "))
Dim Mthn:     Mthn = ErMthn(ErNm)
ErMthLinesByNm = FmtQQ("Private Function ?(?) As String():? = FmtMacro(??):End Function", _
    Mthn, Pm, Mthn, CNm, Calling)
End Function

Private Function ErMthn(ErNm)
ErMthn = "MsgOf_" & ErNm
End Function
