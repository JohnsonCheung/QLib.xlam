Attribute VB_Name = "QDao_F_Schm_SampSchm"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Schm_Samp."
Private Const Asm$ = "QDao"
Public Const StdEleLines$ = _
"E Crt Dte;Req;Dft=Now" & vbCrLf & _
"E Tim Dte" & vbCrLf & _
"E Lng Lng" & vbCrLf & _
"E Mem Mem" & vbCrLf & _
"E Dte Dte" & vbCrLf & _
"E Nm  Txt;Req;Sz=50"
Public Const StdETFLines$ = _
"ETF Nm  * *Nm          " & vbCrLf & _
"ETF Tim * *Tim         " & vbCrLf & _
"ETF Dte * *Dte         " & vbCrLf & _
"ETF Crt * CrtTim       " & vbCrLf & _
"ETF Lng * Si           " & vbCrLf & _
"ETF Mem * Lines *Ft *Fx"

Public Const SpecSchmVbl$ = "Tbl Spec * SpecNm | Lines Ft Si Tim LTimStr CrtTim"
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


