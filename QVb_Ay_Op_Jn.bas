Attribute VB_Name = "QVb_Ay_Op_Jn"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Ay_Op_Jn."
Private Const Asm$ = "QVb"


Function JnSpcApNoBlank$(ParamArray Ap())
Dim Av(): Av = Ap
Stop
'JnSpcApNoBlank = JnCrLf(SyzAyNonBlank(Av))
End Function

Function JnDollarAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnDollarAp = JnDollar(Av)
End Function

Function JnPthSepAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnPthSepAp = JnPthSep(Av)
End Function

Function JnVbarAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnVbarAp = JnVBar(Av)
End Function

Function JnVbarApSpc$(ParamArray Ap())
Dim Av(): Av = Ap
JnVbarApSpc = JnVbarSpc(Av)
End Function

Function JnSpcAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnSpcAp = JnSpc(AyeEmpEle(Av))
End Function


Function JnSemiColonAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnSemiColonAp = JnSemi(AyeEmpEle(Av))
End Function

Private Sub ZZ()
Dim A()
Dim B As Variant

Sy A
End Sub

