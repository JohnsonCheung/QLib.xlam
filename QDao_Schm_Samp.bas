Attribute VB_Name = "QDao_Schm_Samp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Schm_Samp."
Private Const Asm$ = "QDao"
Property Get SampSchm() As String()
Erase XX
X "Tbl LoFmt   *Id Lon"
X "Tbl LoFmtWdt LoFmtId Wdt | Fldss"
X "Tbl LoFmtLvl LoFmtId Lvl | Fldss"
X "Tbl LoFmtBet LoFmtId Fld | FmFld ToFld"
X "Tbl LoFmtTot LoFmtId TotCalc | Fldss"
X "Fld Mem Fldss"
X "Fld Nm  Fld FmFld ToFld"
X "Fld Lng TotCalc"
X "Ele Lvl B Req [VdtRul = >=2 and <=8] Dft=2"
X ""
SampSchm = XX
Erase XX
End Property


