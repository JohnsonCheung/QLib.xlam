Attribute VB_Name = "MXls_Lo_Samp"
Option Explicit
Property Get SampLoVis() As ListObject
Set SampLoVis = LoVis(SampLo)
End Property
Property Get SampLo() As ListObject
Set SampLo = LozRg(RgzSq(SampSqWithHdr, NewA1), "Sample")
End Property

Property Get SampLof() As String()
Erase XX
X "Bet A B C"
X "Lo Nm ABC"
X "Lo Fld A B C D E F G"
X "Lo Hid B C X"
X "Ali Left A B"
X "Ali Right D E"
X "Ali Center F"
X "Wdt 10 A B X"
X "Wdt 20 D C C"
X "Wdt 3000 E F G C"
X "Fmt #,## A B C"
X "Fmt #,##.## D E"
X "Lvl 2 A C"
X "Bdr Left A"
X "Bdr Right G"
X "Bdr Col F"
X "Tot Sum A B"
X "Tot Cnt C"
X "Tot Avg D"
X "Tit A abc | sdf"
X "Tit B abc | sdkf | sdfdf"
X "Cor 12345 A B"
X "Fml F A + B"
X "Fml C A * 2"
X "Lbl A lksd flks dfj"
X "Lbl B lsdkf lksdf klsdj f"
SampLof = XX
Erase XX
End Property

Property Get SampLofTp() As String()
Dim O$()
PushI O, "Lo  Nm     *Nm"
PushI O, "Lo  Fld    *Fld.."
PushI O, "Align Left   *Fld.."
PushI O, "Align Right  *Fld.."
PushI O, "Align Center *Fld.."
PushI O, "Bdr Left   *Fld.."
PushI O, "Bdr Right  *Fld.."
PushI O, "Bdr Col    *Fld.."
PushI O, "Tot Sum    *Fld.."
PushI O, "Tot Avg    *Fld.."
PushI O, "Tot Cnt    *Fld.."
PushI O, "Fmt *Fmt   *Fld.."
PushI O, "Wdt *Wdt   *Fld.."
PushI O, "Lvl *Lvl   *Fld.."
PushI O, "Cor *Cor   *Fld.."
PushI O, "Fml *Fld   *Formula"
PushI O, "Bet *Fld   *Fld1 *Fld2"
PushI O, "Tit *Fld   *Tit"
PushI O, "Lbl *Fld   *Lbl"
SampLofTp = O
End Property


Property Get SampDrzAToJ() As Variant()
Const NC% = 10
Dim J%
For J = 0 To NC - 1
    PushI SampDrzAToJ, Chr(Asc("A") + J)
Next
End Property

Property Get SampSq1() As Variant()
Dim O(), R&, C&
Const NR& = 10000
Const NC& = 1000
ReDim O(1 To NR, 1 To NC)
For R = 1 To NR
For C = 1 To NC
    O(R, C) = R + C
Next
Next
SampSq1 = O
End Property


Property Get SampSqWithHdr() As Variant()
SampSqWithHdr = SqInsDr(SampSq, SampDrzAToJ)
End Property

Property Get SampWs() As Worksheet
Dim O As Worksheet
Set O = NewWs
LozDrs SampDrs, WsRC(O, 2, 2)
Set SampWs = O
WsVis O
End Property
