Attribute VB_Name = "MIde_ConstMth_Val"
Option Explicit
Const CMod$ = "MIde_Gen_Const_ConstVal."
Function ConstValOfFt(ConstNm$)
ConstValOfFt = FtLines(FtzConstQNm(ConstNm))
End Function
Function ConstVal$(ConstQNm$)
Dim Md As CodeModule, ConstNm$
AsgMdAndConstNm Md, ConstNm, _
    ConstQNm
ConstVal = ConstValOfMd(Md, ConstNm)
End Function
Private Sub AsgMdAndConstNm(OMd As CodeModule, OConstNm$, ConstQNm$)

End Sub
Function ConstValOfMd$(Md As CodeModule, ConstNm$)
Dim M$: M = MthLineszMd(Md, "C_" & ConstNm): If M = "" Then Exit Function
If Not IsConstPrp(M) Then Thw CSub, "Not a const method.  It should be [Property Get]", "ConstNm MthLines", ConstNm, M
ConstValOfMd = ConstValOfMth(M)
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
    A = StrBetFstLas(I, """", """")
    B = Replace(A, """""", """")
    PushI O, B
Next
ConstValOfConst = JnCrLf(O)
End Function

Private Function ConstLinesAy(ConstPrpLines$) As String()
Dim Ay$(), O$
O = JnCrLf(O)
Lp:
    Ay = TakP123(O, "Const", vbCrLf & vbCrLf)
    If Si(Ay) = 3 Then
        PushI ConstLinesAy, Ay(1)
        O = Ay(2)
        GoTo Lp
    End If
End Function

Private Sub Z_ConstValOfMth()
Const CSub$ = CMod & "Z_ConstValMthLines"
Dim IsEdt As Boolean, MthLines$, Cas$
GoSub T0
GoSub T1
Exit Sub
T0:
    IsEdt = False
    Cas = "Complex"
    MthLines = TstTxt(CSub, Cas, "MthLines", IsEdt)
    Ept = TstTxt(CSub, Cas, "Ept", IsEdt)
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

Function ConstValOfMd1$(A As CodeModule, ConstNm$)
Dim J%, L$, O$
For J = 1 To A.CountOfDeclarationLines
    L = A.Lines(J, 1)
    O = ConstValOfLinNm(L, ConstNm)
    If O <> "" Then ConstValOfMd1 = O: Exit Function
Next
End Function

Function ConstValOfLinNm$(Lin, ConstNm)
Dim L$: L = RmvMthMdy(Lin)
If Not ShfPfx(L, "Const ") Then Exit Function
If ShfNm(L) <> ConstNm Then Exit Function
If ShfTyChr(L) = "$" Then Thw CSub, "Given constant name is found, but is not a Str", "ConstLin ConstNm", Lin, ConstNm
Dim O$: O = StrBet(L, """", """")
If O = "" Then Thw CSub, "Between DblQuote is nothing", "ConstLin ConstNm", Lin, ConstNm
ConstValOfLinNm = O
End Function

