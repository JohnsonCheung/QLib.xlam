Attribute VB_Name = "MIde_ConstMth_MthLines"
Option Explicit
Const CMod$ = "MIde_Gen_Const_MthLines."

Function ConstPrpLines$(ConstQNm$, IsPub As Boolean)
ConstPrpLines = JnCrLf(ConstPrpLy(ConstQNm, IsPub))
End Function

Private Function ConstPrpLy(ConstQNm$, IsPub As Boolean) As String() 'Ret Ly from ConstPrpFt
Const CSub$ = CMod & "ConstPrpLines"
Dim Ft$: Ft = FtzConstQNm(ConstQNm): If Not HasFfn(Ft$) Then Exit Function
Dim O$()
'    PushI O, IIf(IsPub, "", "Private ") & "Property Get " & ConstNm & "() As String()"
    Dim L, Fst As Boolean: Fst = True
    For Each L In Itr(LyzFt(Ft$))
        PushIAy O, CdLyzPushStr(L, Fst)
        If Fst Then
            Fst = False
        End If
    Next
    PushI O, "End Property"
'ConstPrpLines = O
End Function

Private Sub Z_ExprLyzStr()
Brw ExprLyzStr(StrzCurPjf)
End Sub

Private Function CdLyzPushStr(S, ByVal Fst As Boolean) As String()
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
CdLyzPushStr = O
End Function

Private Sub Z_ConstPrpLines()
Const TstId% = 2
Const CSub$ = CMod & "Z_ConstPrpLines"
GoSub ZZ
Exit Sub
GoSub T0
GoSub T1
GoSub T2
Exit Sub
'--
Dim Nm$, ConstVal$, IsPub As Boolean
Dim IsEdt As Boolean, Cas$
T0:
    Cas = "Complex1"
    IsEdt = False
    Nm = "ZZ_B"
    ConstVal = TstTxt(TstId, CSub, Cas, "ConstVal", IsEdt:=False)
    Ept = TstTxt(TstId, CSub, "Complex1", "Ept", IsEdt)
    IsPub = True
    GoTo Tst

T1:
    IsEdt = True
    ConstVal = MthLinesByMdMth(CurMd, "Chunk")
    BrwStr ConstVal
    Stop
    Nm = "ZZ_A"
    IsPub = True
    Ept = TstTxt(TstId, CSub, "Complex", "Ept", IsEdt)
    GoTo Tst

T2:
    IsEdt = False
    Nm = "ZZ_A"
    ConstVal = "AAA"
    Ept = JnCrLf(Array("", _
        "Private Function ZZ_A$()", _
        "Const A_1$ = ""AAA""", _
        "", _
        "ZZ_A = A_1", _
        "End Function"))
    GoTo Tst
Tst:
    If IsEdt Then Return
    If ConstVal = "" Then Stop
'    Act = ConstPrpLines(ConstVal, Nm, IsPub)
    Brw Act: Stop
    C
    ShwTstOk CSub, Cas
    Return
ZZ:
    Dim V$: V = JnCrLf(AywFstNEle(SrczPj(CurPj), 5000))
    Stop
'    Brw ConstPrpLines("AA", V, IsPub:=True)
    Return
End Sub

Private Sub ZZ()
Dim A$
Dim B As Boolean
End Sub

Private Sub Z()
Z_ConstPrpLines
End Sub
Private Property Get C_A$()
Const A_1$ = "sldkfj skldjf slkdfj sd" & _
vbCrLf & "sdfkljsdf" & _
vbCrLf & ""

C_A = A_1
End Property

