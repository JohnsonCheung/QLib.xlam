Attribute VB_Name = "MXls_Lo_Lof"
Option Explicit
Public Const LofKK$ = "Ali Bdr Bet Cor Fml Fmt Lbl Lvl Tit Tot Wdt"
Function Lof(LofDefDic As Dictionary) As Lof
Set Lof = New Lof
Lof.Init Ali, Bdr, Bet, Cor, Fml, Fmt, Lbl, Lvl, Tit, Tot, Wdt
End Function
Private Function Wdt() As LofWdt()
End Function
Private Function Cor() As LofCor()
End Function
Private Function Lvl() As LofLvl()
End Function
Private Function Fmt() As LofFmt()
End Function
Private Function Ali() As LofAli()
End Function
Private Function Tot() As LofTot()
End Function
Private Function Bdr() As LofBdr()
End Function
Private Function Fml() As LofFml()
Dim U%: U = UB(Fml)
If U < 0 Then Exit Function
Dim Fld$, X$, J%, L, O() As LofFml
ReDim O(U)
For Each L In Fml
    AsgTRst L, Fld, X
    O(J).Fld = Fld
    O(J).Fml = X
    J = J + 1
Next
Fml = O
End Function
Private Function Lbl() As LofLbl()
Dim U%: U = UB(Lbl)
If U < 0 Then Exit Function
Dim Fld$, X$, J%, L, O() As LofLbl
ReDim O(U)
For Each L In Lbl
    AsgTRst L, Fld, X
    O(J).Fld = Fld
    O(J).Lbl = X
    J = J + 1
Next
Lbl = O
End Function
Private Function Tit() As LofTit()
Dim U%: U = UB(Tit)
If U < 0 Then Exit Function
Dim Fld$, X$, J%, L, O() As LofTit
ReDim O(U)
For Each L In Tit
    AsgTRst L, Fld, X
    O(J).Fld = Fld
    O(J).Tit = X
    J = J + 1
Next
Tit = O
End Function
Private Function Bet() As LofBet()
Dim U%: U = UB(Bet)
If U < 0 Then Exit Function
Dim Fld$, FmFld$, ToFld$, SumFld$, J%, L, O() As LofBet
ReDim O(U)
For Each L In Bet
    Asg2TRst L, SumFld, FmFld, ToFld
    Set O(J) = New LofBet
    O(J).Init Fld, FmFld, ToFld
    J = J + 1
Next
Bet = O
End Function
Sub BrwSampLoFmtrTp()
BrwStr SampLoFmtrTp
End Sub

