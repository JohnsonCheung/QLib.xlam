Attribute VB_Name = "QXls_Samp"
Option Explicit
Private Const CMod$ = "BSamp."
Property Get SampLoVis() As ListObject
Set SampLoVis = ShwLo(SampLo)
End Property
Property Get SampLo() As ListObject
Set SampLo = CrtLozRg(RgzSq(SampSqWithHdr, NewA1), "Sample")
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


Property Get SampDrOfAToJ() As Variant()
Const NC% = 10
Dim J%
For J = 0 To NC - 1
    PushI SampDrOfAToJ, Chr(Asc("A") + J)
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
SampSqWithHdr = InsSqr(SampSq, SampDrOfAToJ)
End Property

Property Get SampWs() As Worksheet
Dim O As Worksheet
Set O = NewWs
LozDrs SampDrs, WsRC(O, 2, 2)
Set SampWs = O
ShwWs O
End Property
