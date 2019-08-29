Attribute VB_Name = "QVb_Ay_Op_AyJn"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Op_Jn."
Private Const Asm$ = "QVb"


Function JnSpcApNoBlnk$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
Stop
'JnSpcApNoBlnk = JnCrLf(SyzAyNB(Av))
End Function

Function JnDollarAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnDollarAp = JnDollar(Av)
End Function

Function JnPthSepAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnPthSepAp = JnPthSep(Av)
End Function

Function JnVbarAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnVbarAp = JnVBar(Av)
End Function

Function JnVbarApSpc$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnVbarApSpc = JnVbarSpc(Av)
End Function

Function JnSpcAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnSpcAp = JnSpc(AeEmpEle(Av))
End Function


Function JnSemiColonAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnSemiColonAp = JnSemi(AeEmpEle(Av))
End Function

Private Sub Z()
Dim A()
Dim B As Variant

Sy A
End Sub


'
