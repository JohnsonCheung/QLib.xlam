Attribute VB_Name = "QVb_Str_Term"
Option Explicit
Private Const CMod$ = "MTerm."
Const NS$ = "Vb.Str.Term"
Private Const Asm$ = "Q"
Function RmvTerm$(Lin, Term$())
RmvTerm = JnTerm(MinusAy(TermAy(Lin), Term))
End Function
Function TermLin(TermAy)
TermLin = TLin(TermAy)
End Function
Function TLin(TermAy)
TLin = JnTerm(TermAy)
End Function

Function TLinzAp$(ParamArray TermAp())
Dim Av(): Av = TermAp
TLinzAp = JnTerm(Av)
End Function

Function JnTerm$(TermAy)
JnTerm = JnSpc(QuoteSqzAyIf(RmvBlankzAy(TermAy)))
End Function

Function LinzTermAy$(TermAy)
LinzTermAy = TLin(TermAy)
End Function

Function TermAset(S) As Aset
Set TermAset = AsetzAy(TermAy(S))
End Function

Function TermItr(S)
Asg Itr(TermAy(S)), TermItr
End Function
Function TermAyzDr(Dr()) As String()

End Function
Function TermAy(Lin) As String()
Dim L$, J%
L = Lin
While L <> ""
    J = J + 1: If J > 5000 Then Stop
    PushNonBlank TermAy, ShfT1(L)
Wend
End Function
Function ShfTerm(OLin$, X$) As Boolean
If T1(OLin) = X Then
    ShfTerm = True
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


Private Sub ZZ()
Z_ShfT1
MVb_Lin_Term:
End Sub
