Attribute VB_Name = "MTerm"
Option Explicit
Const CMod$ = "MVb_Lin_Term."
Const Ns$ = "Vb.Str.Term"
Const Asm$ = "QVb"
Function RmvTermAy$(S$, TermAy$())
Dim T$, I
T = T1(S)
For Each I In TermAy
    If I = T Then
        RmvTermAy = LTrim(Mid(LTrim(S), Len(T) + 1))
        Exit Function
    End If
Next
RmvTermAy = S
End Function
Function TermLin$(TermAy$())
TermLin = TLin(TermAy)
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
JnTermAy = JnSpc(SyQuoteSqIf(SyRmvBlank(TermAy)))
End Function

Function TermAyzTT(TT$) As String()
Const CSub$ = CMod & "TermAyzTT"
Dim T
For Each T In TermItr(TT)
    PushI TermAyzTT, T
Next
End Function

Function LinzTermAy$(TermAy$())
LinzTermAy = TLin(TermAy)
End Function

Function TermAset(S$) As Aset
Set TermAset = AsetzAy(TermAy(S))
End Function

Function TermItr(S$)
Asg Itr(TermAy(S)), TermItr
End Function
Function TermAyzDr(Dr()) As String()

End Function
Function TermAy(Lin$) As String()
Dim L$, J%
L = Lin
While L <> ""
    J = J + 1: If J > 5000 Then Stop
    PushNonBlankStr TermAy, ShfT1(L)
Wend
End Function
Function ShfTermX(OLin$, X$) As Boolean
If T1(OLin) = X Then
    ShfTermX = True
    OLin = RmvT1(OLin)
End If
End Function

Function ShfT1$(OLin$)
ShfT1 = T1(OLin)
OLin = RmvT1(OLin)
End Function

Function ShfTermDot$(OLin$)
With Brk2Dot(OLin, NoTrim:=True)
    ShfTermDot = .S1
    OLin = .S2
End With
End Function

Private Sub Z_ShfT1()
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
    Act = ShfT1(O)
    C
    Ass O = OEpt
    Return
End Sub


Private Sub Z()
Z_ShfT1
MVb_Lin_Term:
End Sub

