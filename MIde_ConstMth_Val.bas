Attribute VB_Name = "MIde_ConstMth_Val"
Option Explicit
Const CMod$ = "MIde_Gen_Const_ConstVal."
Function ConstValOfFt(ConstNm$)
ConstValOfFt = LineszFt(FtzConstQNm(ConstNm))
End Function
Function ConstVal$(ConstQNm$)
Dim Md As CodeModule, ConstNm$
AsgMdAndConstNm Md, ConstNm, _
    ConstQNm
ConstVal = ConstValByMd(Md, ConstNm)
End Function
Private Sub AsgMdAndConstNm(OMd As CodeModule, OConstNm$, ConstQNm$)

End Sub
Function ConstValByMd$(Md As CodeModule, ConstNm$)
Dim M$: M = MthLinesByMdMth(Md, "C_" & ConstNm): If M = "" Then Exit Function
If Not IsConstPrp(M) Then Thw CSub, "Not a const method.  It should be [Property Get]", "ConstNm MthLines", ConstNm, M
ConstValByMd = ConstValOfMth(M)
End Function

Private Function IsConstPrp(MthLines$) As Boolean
Dim A As MthNm3: Set A = MthNm3(FstLin(MthLines))
If A.MthTy = "Property Get" Then Exit Function
If Not HasPfx(A.Nm, "C_") Then Exit Function
IsConstPrp = True
End Function
Function ConstValOfMth$(MthLines$)
Dim O$(), ConstLines
For Each ConstLines In Itr(ConstLinesAy(MthLines))
    PushI O, ConstValOfConst(ConstLines)
Next
ConstValOfMth = Jn(O)
End Function

Private Function ConstValOfConst$(C)
Dim I, O$(), A$, B$
For Each I In SplitCrLf(C)
    A = BetFstLas(I, """", """")
    B = Replace(A, """""", """")
    PushI O, B
Next
ConstValOfConst = JnCrLf(O)
End Function

Private Function ConstLinesAy(ConstPrpLines$) As String()
Dim Ay$(), O$
O = JnCrLf(O)
Lp:
    Ay = P123(O, "Const", vbCrLf & vbCrLf)
    If Si(Ay) = 3 Then
        PushI ConstLinesAy, Ay(1)
        O = Ay(2)
        GoTo Lp
    End If
End Function

Private Sub Z_ConstValOfMth()
Const TstId% = 3
Const CSub$ = CMod & "Z_ConstValMthLines"
Dim MthLines$, Cas$, IsEdt As Boolean
GoSub T0
GoSub T1
Exit Sub
T0:
    IsEdt = False
    Cas = "Complex"
    MthLines = TstTxt(TstId, CSub, Cas, "MthLines", IsEdt:=True)
    Ept = TstTxt(TstId, CSub, Cas, "Ept", IsEdt)
    If IsEdt Then Return
    GoTo Tst
T1:
   
    Return
Tst:
    Act = ConstValOfMth(MthLines)
    Brw Act
    Stop
    C
    Return
End Sub

Private Sub Z()
Z_ConstValOfMth
MIde_Gen_Const_ConstVal:
End Sub

Function ConstValByMd1$(A As CodeModule, ConstNm$)
Dim J%, L$, O$
For J = 1 To A.CountOfDeclarationLines
    L = A.Lines(J, 1)
    O = ConstValOfLinNm(L, ConstNm)
    If O <> "" Then ConstValByMd1 = O: Exit Function
Next
End Function

Function ConstValOfLinNm$(Lin, ConstNm)
Dim L$: L = RmvMthMdy(Lin)
If Not ShfPfx(L, "Const ") Then Exit Function
If ShfNm(L) <> ConstNm Then Exit Function
If ShfTyChr(L) = "$" Then Thw CSub, "Given constant name is found, but is not a Str", "ConstLin ConstNm", Lin, ConstNm
Dim O$: O = Bet(L, """", """")
If O = "" Then Thw CSub, "Between DblQuote is nothing", "ConstLin ConstNm", Lin, ConstNm
ConstValOfLinNm = O
End Function

