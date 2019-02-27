Attribute VB_Name = "MIde_ConstMth_Val"
Option Explicit
Const CMod$ = "MIde_Gen_Const_ConstVal."
Function ConstValzFt(ConstNm$)
ConstValzFt = FtLines(FtzConstQNm(ConstNm))
End Function
Function ConstVal$(ConstQNm$)

End Function
Function ConstValzMd$(Md As CodeModule, ConstNm$)
Dim M$: M = MthLineszMd(Md, "C_" & ConstNm): If M = "" Then Exit Function
If Not IsConstMth(M) Then Thw CSub, "Not a const method.  It should be [Property Get]", "ConstNm MthLines", ConstNm, M
ConstValzMd = ConstValzMth(M)
End Function

Private Function IsConstMth(MthLines$) As Boolean
Dim A As MthNm3: Set A = MthNm3(FstLin(MthLines))
If A.MthTy = "Property Get" Then Exit Function
If Not HasPfx(A.Nm, "C_") Then Exit Function
IsConstMth = True
End Function
Function ConstValzMth$(MthLines$)
Dim O$(), ConstLines
For Each ConstLines In Itr(ConstLinesAy(MthLines))
    PushI O, ConstValzConst(ConstLines)
Next
ConstValzMth = Jn(O)
End Function

Private Function ConstValzConst$(C)
Dim I, O$(), A$, B$
For Each I In SplitCrLf(C)
    A = TakBetFstLas(I, """", """")
    B = Replace(A, """""", """")
    PushI O, B
Next
ConstValzConst = JnCrLf(O)
End Function

Private Function ConstLinesAy(ConstMthLines$) As String()
Dim Ay$(), O$
O = JnCrLf(O)
Lp:
    Ay = TakP123(O, "Const", vbCrLf & vbCrLf)
    If Sz(Ay) = 3 Then
        PushI ConstLinesAy, Ay(1)
        O = Ay(2)
        GoTo Lp
    End If
End Function

Private Sub Z_ConstValzMth()
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
    Act = ConstValzMth(MthLines)
    Brw Act
    Stop
    C
    Return
End Sub

Private Sub Z()
Z_ConstValzMth
MIde_Gen_Const_ConstVal:
End Sub

Function ConstValzMd1$(A As CodeModule, ConstNm$)
Dim J%, L$, O$
For J = 1 To A.CountOfDeclarationLines
    L = A.Lines(J, 1)
    O = ConstValzLinNm(L, ConstNm)
    If O <> "" Then ConstValzMd1 = O: Exit Function
Next
End Function

Function ConstValzLinNm(Lin, ConstNm) As String
Dim L$: L = RmvMthMdy(Lin)
If Not ShfX(L, "Const ") Then Exit Function
If ShfNm(L) <> ConstNm Then Exit Function
If ShfMthChr(L) = "$" Then Thw CSub, "Given constant name is found, but is not a Str", "ConstLin ConstNm", Lin, ConstNm
Dim O$: O = TakBet(L, """", """")
If O = "" Then Thw CSub, "Between DblQuote is nothing", "ConstLin ConstNm", Lin, ConstNm
ConstValzLinNm = O
End Function

