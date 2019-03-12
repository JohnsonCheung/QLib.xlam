Attribute VB_Name = "MVb_Lin_Term"
Option Explicit
Const CMod$ = "MVb_Lin_Term."
Function RmvTermAy$(Lin, Ay$())
Dim T$, I
T = T1(Lin)
For Each I In Ay
    If I = T Then
        RmvTermAy = LTrim(Mid(LTrim(Lin), Len(T) + 1))
        Exit Function
    End If
Next
RmvTermAy = Lin
End Function

Function TLin$(TermAy$())
TLin = JnTermAy(TermAy)
End Function

Function TLinzAp$(ParamArray TermAp())
Dim Av(): Av = TermAp
TLinzAp = JnTermAy(SyzAv(Av))
End Function

Function JnTermAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnTermAp = JnTermAy(SyzAv(Av))
End Function

Function JnTermAy$(TermAy$())
JnTermAy = JnSpc(AyQuoteSqIf(SyRmvBlank(TermAy)))
End Function

Function TermAyzTT(TT) As String()
Const CSub$ = CMod & "TermAyzTT"
Dim T
For Each T In TermItr(TT)
    PushI TermAyzTT, T
Next
Select Case True
Case IsStr(TT): TermAyzTT = TermAy(TT)
Case IsSy(TT): TermAyzTT = TT
Case Else: Thw CSub, "Given TT must be Str or Sy", "TypeName TT", TypeName(TT), TT
End Select
End Function

Function LinzTermAy$(TermAy)
LinzTermAy = JnSpc(AyQuoteSqIf(TermAy))
End Function

Function TermAset(Lin) As Aset
Set TermAset = AsetzAy(TermAy(Lin))
End Function

Function TermItr(NN)
Asg TermAyzNN(NN), TermItr
End Function

Function CvNy(Ny0) As String()
Const CSub$ = CMod & "CvNy"
Select Case True
Case IsMissing(Ny0) Or IsEmpty(Ny0)
Case IsStr(Ny0): CvNy = TermAy(Ny0)
Case IsSy(Ny0): CvNy = Ny0
Case IsArray(Ny0): CvNy = SyzAy(Ny0)
Case Else: Thw CSub, "Given Ny0 must be Missing | Empty | Str | Sy | Ay", "TypeName-Ny0", TypeName(Ny0)
End Select
End Function

Function TermAyzNN(NN) As String()
Select Case True
Case IsStr(NN): TermAyzNN = TermAy(NN)
Case IsSy(NN): TermAyzNN = NN
Case Else: Thw CSub, "NN must be String or Sy", "TypeName(NN)", TypeName(NN)
End Select
End Function
Function TermAy(Lin) As String()
Dim L$, J%
L = Lin
While L <> ""
    J = J + 1: If J > 5000 Then Stop
    PushNonBlankStr TermAy, ShfTerm(L)
Wend
End Function

Function ShfT$(O)
ShfT = ShfTerm(O)
End Function

Function ShfX(O, X$) As Boolean
If T1(O) = X Then
    ShfX = True
    O = RmvT1(O)
End If
End Function

Private Function ShfTerm1$(O)
Dim A$
AsgAp BrkBkt(O, "["), A, ShfTerm1, O
End Function

Function ShfTerm$(O)
Dim A$
    A = LTrim(O)
If FstChr(A) = "[" Then ShfTerm = ShfTerm1(O): Exit Function
Dim P%
    P = InStr(A, " ")
If P = 0 Then
    ShfTerm = A
    O = ""
    Exit Function
End If
ShfTerm = Left(A, P - 1)
O = LTrim(Mid(A, P + 1))
End Function

Private Sub Z_ShfT()
Dim O$, OEpt$
O = " S   DFKDF SLDF  "
OEpt = "DFKDF SLDF  "
Ept = "S"
GoSub Tst
'
O = " AA BB "
Ept = "AA"
OEpt = "BB "
GoSub Tst
'
Exit Sub
Tst:
    Act = ShfT(O)
    C
    Ass O = OEpt
    Return
End Sub


Private Sub Z()
Z_ShfT
MVb_Lin_Term:
End Sub

