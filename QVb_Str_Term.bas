Attribute VB_Name = "QVb_Str_Term"
Option Explicit
Private Const CMod$ = "MTerm."
Const Ns$ = "Vb.Str.Term"
Private Const Asm$ = "Q"
Function RmvTermSy$(S$, TermSy$())
Dim T$, I
T = T1(S)
For Each I In TermSy
    If I = T Then
        RmvTermSy = LTrim(Mid(LTrim(S), Len(T) + 1))
        Exit Function
    End If
Next
RmvTermSy = S
End Function
Function TermLin$(TermSy$())
TermLin = TLin(TermSy)
End Function
Function TLin$(TermSy$())
TLin = JnTermSy(TermSy)
End Function

Function TLinzAp$(ParamArray TermAp())
Dim Av(): Av = TermAp
TLinzAp = JnTermSy(SyzAv(Av))
End Function

Function JnTermAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnTermAp = JnTermSy(SyzAv(Av))
End Function

Function JnTermSy$(TermSy$())
JnTermSy = JnSpc(SyQuoteSqIf(SyRmvBlank(TermSy)))
End Function

Function LinzTermSy$(TermSy$())
LinzTermSy = TLin(TermSy)
End Function

Function TermAset(S$) As Aset
Set TermAset = AsetzAy(TermSy(S))
End Function

Function TermItr(S$)
Asg Itr(TermSy(S)), TermItr
End Function
Function TermSyzDr(Dr()) As String()

End Function
Function TermSy(Lin$) As String()
Dim L$, J%
L = Lin
While L <> ""
    J = J + 1: If J > 5000 Then Stop
    PushNonBlank TermSy, ShfT1(L)
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

