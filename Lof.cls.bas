VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Ali() As LofAli
Private A_Bdr() As LofBdr
Private A_Bet() As LofBet
Private A_Cor() As LofCor
Private A_Fml() As LofFml
Private A_Fmt() As LofFmt
Private A_Lbl() As LofLbl
Private A_Lvl() As LofLvl
Private A_Tit() As LofTit
Private A_Tot() As LofTot
Private A_Wdt() As LofWdt

Friend Function Init( _
 Ali() As LofAli, _
 Bdr() As LofBdr, _
 Bet() As LofBet, _
 Cor() As LofCor, _
 Fml() As LofFml, _
 Fmt() As LofFmt, _
 Lbl() As LofLbl, _
 Lvl() As LofLvl, _
 Tit() As LofTit, _
 Tot() As LofTot, _
 Wdt() As LofWdt _
) As Lof
A_Ali = Ali
A_Bdr = Bdr
A_Bet = Bet
A_Cor = Cor
A_Fml = Fml
A_Fmt = Fmt
A_Lbl = Lbl
A_Lvl = Lvl
A_Tit = Tit
A_Tot = Tot
A_Wdt = Wdt
Set Init = Me
End Function
Function Ali() As LofAli()
Ali = A_Ali
End Function
Function Bdr() As LofBdr()
Bdr = A_Bdr
End Function
Function Bet() As LofBet()
Bet = A_Bet
End Function
Function Cor() As LofCor()
Cor = A_Cor
End Function
Function Fml() As LofFml()
Fml = A_Fml
End Function
Function Fmt() As LofFmt()
Fmt = A_Fmt
End Function
Function Lbl() As LofLbl()
Lbl = A_Lbl
End Function
Function Lvl() As LofLvl()
Lvl = A_Lvl
End Function
Function Tit() As LofTit()
Tit = A_Tit
End Function
Function Tot() As LofTot()
Tot = A_Tot
End Function
Function Wdt() As LofWdt()
Wdt = A_Wdt
End Function
