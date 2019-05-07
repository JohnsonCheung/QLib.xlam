Attribute VB_Name = "QIde_Parse_Syc"
Option Explicit
Private Const Asm$ = "QIde"
Private Const CMod$ = "MIde_ConstMth_Val."
Const DoczSyc$ = "Sy-Const.  Each Module/Class may have some ro prp [C_{SycNm}] of type String()"
Const DoczSycNm$ = "Sy-Const-Nm.  It is a name without pfx C_"
Const DoczSycMNm$ = "Sy-Const-Mth-Nm.  It is a name with pfx C_, which is `C_{SycNm}"
Const DoczSycQNm$ = "Sy-Const-Qualitied-Nm.  It is a `{CmpNm}.{SycNm}`"
Const DoczSycMLines$ = "Sy-Const-Mth-Lines."
Const DoczSycFt$ = "Sy-Const-Ft.  It comes from SycNm."
Private Type MdSyc
    Md As CodeModule
    SycNm As String
End Type
Function SycFt$(SycQNm$)
SycFt = SycHom & SycQNm & ".txt"
End Function
Property Get SycHom$()
Static X$: If X = "" Then X = AddFdrEns(TmpHom & "Syc\")
SycHom = X
End Property
Function SycValzFt(SycNm$) As String()
SycValzFt = LyzFt(FtzCnstQNm(SycNm))
End Function
Function SycVal(A As MdSyc) As String()
SycVal = SycValzMdSyc(MdSyc(CnstQNm))
End Function
Function MdzSycNm(SycNm$) As CodeModule
Dim A$(): A = MdNyzSycNm(SycNm)
Select Case Si(A)
Case 0: Thw CSub, "SycNm not in any Md", "SycNm", SycNm
Case 1: Set MdSyc = Md(A(0))
Case Else: Thw CSub, "SycNm is found in more than one module", "SycNm MdNy", SycNm, A
End Select
End Function
Private Function MdSyc(CnstQNm$) As MdSyc
Dim O As MdSyc
With Brk2Dot(CnstQNm)
    If .S1 = "" Then
        Set O.Md = MdSyc(.S2)
    Else
        Set O.Md = Md(.S1)
    End If
    O.SycNm = .S2
End With
End Function

Private Sub Z_MdNyzSycNm()
D MdNyzSycNm("CMod")
End Sub

Function MdNyzSycNm(SycNm$) As String()
MdNyzSycNm = MdNyzSycNmPj(CurPj, SycNm)
End Function

Function MdNyzSycNmPj(Pj As VBProject, SycNm$) As String()
Dim C As VBComponent
For Each C In Pj.VBComponents
    If HasSycNm(C.CodeModule, SycNm) Then
        PushI MdNyzSycNmPj, C.Name
    End If
Next
End Function
Function CnstBrkzMd$(Md As CodeModule, SycNm$)
Dim M$: M = MthLinesByMdMth(Md, "C_" & SycNm): If M = "" Then Exit Function
If Not IsMthLinzSyc(M) Then Thw CSub, "Not a const method.  It should be [Property Get]", "SycNm MthLines", SycNm, M
CnstBrkzMd = SycVal(M)
End Function

Private Function IsMthLinzSyc(MthLin$) As Boolean
Dim A As MthNm3: Set A = MthNm3(FstLin(MthLin))
If A.MthTy = "Property Get" Then Exit Function
If Not HasPfx(A.Nm, "C_") Then Exit Function
IsMthLinzSyc = True
End Function
Function SycValzSycMLines(MthLines$) As String()
Dim XLinSy$(): XLinSy = RmvPfxzSy(SywPfx(SplitCrLf(MthLines), "X """), "X ")
SycValzSycMLines = TakVbStrzSy(XLinSy)
End Function
Function TakVbStr$(VbStr$)
If FstChr(VbStr) <> """" Then Thw CSub, "FstChr of VbStr must be DblQuote", "VbStr", VbStr
Dim P%: P = InStr(2, VbStr, """")
If P = 0 Then Thw CSub, "There is no ending DblQuote", "VbStr", VbStr
TakVbStr = Mid(VbStr, 2, P - 2)
End Function
Function TakVbStrzSy(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI TakVbStrzSy, TakVbStr(CStr(I))
Next
End Function
Private Function CnstBrkzConst$(C)
Dim I, O$(), A$, B$
For Each I In SplitCrLf(C)
    A = BetFstLas(I, """", """")
    B = Replace(A, """""", """")
    PushI O, B
Next
CnstBrkzConst = JnCrLf(O)
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

Private Sub Z_SycVal()
Const TstId& = 3
Const CSub$ = CMod & "Z_CnstBrkMthLines"
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
    Act = SycVal(MthLines)
    Brw Act
    Stop
    C
    Return
End Sub

Private Sub Z()
Z_SycVal
MIde_Gen_Const_CnstBrk:
End Sub

Function CnstBrkzMd1$(A As CodeModule, SycNm$)
Dim J%, L$, O$
For J = 1 To A.CountOfDeclarationLines
    L = A.Lines(J, 1)
    O = CnstBrkzLinNm(L, SycNm)
    If O <> "" Then CnstBrkzMd1 = O: Exit Function
Next
End Function

Function CnstBrkzLinNm$(Lin, SycNm)
Dim L$: L = RmvMthMdy(Lin)
If Not ShfPfx(L, "Const ") Then Exit Function
If ShfNm(L) <> SycNm Then Exit Function
If ShfTyChr(L) = "$" Then Thw CSub, "Given constant name is found, but is not a Str", "ConstLin SycNm", Lin, SycNm
Dim O$: O = Bet(L, """", """")
If O = "" Then Thw CSub, "Between DblQuote is nothing", "ConstLin SycNm", Lin, SycNm
CnstBrkzLinNm = O
End Function

