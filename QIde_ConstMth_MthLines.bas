Attribute VB_Name = "QIde_ConstMth_MthLines"
Option Compare Text
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_ConstMth_MthL."

Function ConstPrpLines$(CnstQNm$, IsPub As Boolean)
ConstPrpLines = JnCrLf(ConstPrpLy(CnstQNm, IsPub))
End Function

Private Function ConstPrpLy(CnstQNm$, IsPub As Boolean) As String() 'Ret Ly from ConstPrpFt
Const CSub$ = CMod & "ConstPrpLines"
Dim Ft$: Ft = FtzCnstQNm(CnstQNm): If Not HasFfn(Ft) Then Exit Function
Dim O$()
'    PushI O, IIf(IsPub, "", "Private ") & "Property Get " & Cnstn & "() As String()"
    Dim L, Fst As Boolean: Fst = True
    For Each L In Itr(LyzFt(Ft))
        PushIAy O, CdLyzPushItr(L, Fst)
        If Fst Then
            Fst = False
        End If
    Next
    PushI O, "End Property"
'ConstPrpLines = O
End Function

Private Sub Z_ExprLyzStr()
'Brw ExprLyzStr(StrOfPjfP)
End Sub

Private Function CdLyzPushItr(S, ByVal Fst As Boolean) As String()
Dim CdLin, LasL%, O$()
Dim CdLy$(): CdLy = ExprLyzStr(S)
LasL = Si(CdLy)
Dim J%
For Each CdLin In Itr(CdLy(S))
    Select Case True
    Case Fst:      PushI O, "X2 " & CdLin: Fst = False
    Case LasL = J: PushI O, "X " & CdLin
    Case Else:     PushI O, "X1 " & CdLin
    End Select
    J = J + 1
Next
CdLyzPushItr = O
End Function

Private Sub Z_ConstPrpLines()
Const TstId& = 2
Const CSub$ = CMod & "Z_ConstPrpLines"
GoSub Z
Exit Sub
GoSub T0
GoSub T1
GoSub T2
Exit Sub
'--
Dim Nm$, CnstBrk$, IsPub As Boolean
Dim IsEdt As Boolean, Cas$
T0:
    Cas = "Complex1"
    IsEdt = False
    Nm = "Z_B"
    CnstBrk = TstTxt(TstId, CSub, Cas, "CnstBrk", IsEdt:=False)
    Ept = TstTxt(TstId, CSub, "Complex1", "Ept", IsEdt)
    IsPub = True
    GoTo Tst

T1:
    IsEdt = True
    'CnstBrk = MthLzNmzMd(CMd, "Chunk")
    BrwStr CnstBrk
    Stop
    Nm = "Z_A"
    IsPub = True
    Ept = TstTxt(TstId, CSub, "Complex", "Ept", IsEdt)
    GoTo Tst

T2:
    IsEdt = False
    Nm = "Z_A"
    CnstBrk = "AAA"
    Ept = JnCrLf(Array("", _
        "Private Function Z_A$()", _
        "Const A_1$ = ""AAA""", _
        "", _
        "Z_A = A_1", _
        "End Function"))
    GoTo Tst
Tst:
    If IsEdt Then Return
    If CnstBrk = "" Then Stop
'    Act = ConstPrpLines(CnstBrk, Nm, IsPub)
    Brw Act: Stop
    C
    ShwTstOk CSub, Cas
    Return
Z:
    Dim V$: V = JnCrLf(FstNEle(SrczP(CPj), 5000))
    Stop
'    Brw ConstPrpLines("AA", V, IsPub:=True)
    Return
End Sub

Private Sub Z()
Dim A$
Dim B As Boolean
End Sub

Private Property Get C_A$()
Const A_1$ = "sldkfj skldjf slkdfj sd" & _
vbCrLf & "sdfkljsdf" & _
vbCrLf & ""

C_A = A_1
End Property

