Attribute VB_Name = "QVb_Dta_Re"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Re."
Private Const Asm$ = "QVb"

Private Sub Z_ReMatch()
Dim A As MatchCollection
Dim R  As RegExp: Set R = RegExp("m[ae]n")
Set A = R.Execute("alskdflfmEnsdklf")
Stop
End Sub

Private Sub Z_ReRpl()
Dim R As RegExp: Set R = RegExp("(.+)(m[ae]n)(.+)")
Dim Act$: Act = R.Replace("a men is male", "$1male$3")
Ass Act = "a male is male"
End Sub

'
