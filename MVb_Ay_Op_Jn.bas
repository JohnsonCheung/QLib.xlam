Attribute VB_Name = "MVb_Ay_Op_Jn"
Option Explicit


Function JnSpcApNoBlank$(ParamArray Ap())
Dim Av(): Av = Ap
JnSpcApNoBlank = JnCrLf(SyzAyNonBlank(Av))
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
JnVbarAp = JnVbar(Av)
End Function

Function JnVbarApSpc$(ParamArray Ap())
Dim Av(): Av = Ap
JnVbarApSpc = JnVbarSpc(Av)
End Function

Function JnSpcAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnSpcAp = JnSpc(AyeEmpEle(Av))
End Function

Function JnSpcApes$(ParamArray Ap())
Dim Av(): Av = Ap
JnSpcApes = JnCrLf(AyeEmpEle(Av))
End Function

Function ApScl$(ParamArray Ap())
Dim Av(): Av = Ap
ApScl = JnSemi(AyeEmpEle(Av))
End Function

Private Sub ZZ()
Dim A()
Dim B As Variant

Sy A
End Sub

Private Sub Z()
End Sub
