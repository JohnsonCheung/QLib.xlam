Attribute VB_Name = "QIde_DoMth_AddMthCol"
Option Explicit
Option Compare Text

Function AddColoHasPm(WiMthLin As Drs) As Drs
'Ret         : Add col-HasPm
Dim I%: I = IxzAy(WiMthLin.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In Itr(WiMthLin.Dy)
    Dim MthLin$: MthLin = Dr(I)
    Dim HasPm As Boolean: HasPm = BetBkt(MthLin) <> ""
    PushI Dr, HasPm
    PushI Dy, Dr
Next
AddColoHasPm = AddColzFFDy(WiMthLin, "HasPm", Dy)
End Function

Function AddColoTyChr(WiMthLin As Drs) As Drs
'Ret         : Add col-HasPm
Dim I%: I = IxzAy(WiMthLin.Fny, "MthLin")
Dim Dr, Dy(): For Each Dr In Itr(WiMthLin.Dy)
    Dim MthLin$: MthLin = Dr(I)
    Dim TyChr$: TyChr = MthTyChr(MthLin)
    PushI Dr, TyChr
    PushI Dy, Dr
Next
AddColoTyChr = AddColzFFDy(WiMthLin, "TyChr", Dy)
End Function

Function AddColoShtPm(WiMthPm As Drs) As Drs
'Ret         : Add col-ShtPm
Dim I%: I = IxzAy(WiMthPm.Fny, "MthPm")
Dim Dr, Dy(): For Each Dr In Itr(WiMthPm.Dy)
    Dim MthPm$: MthPm = Dr(I)
    Dim ShtPm1$: ShtPm1 = ShtPm(MthPm)
    PushI Dr, ShtPm1
    PushI Dy, Dr
Next
AddColoShtPm = AddColzFFDy(WiMthPm, "ShtPm", Dy)
End Function

Function AddColoMthPm(WiMthLin As Drs, Optional IsDrp As Boolean) As Drs
AddColoMthPm = AddColzBetBkt(WiMthLin, "MthLin:MthPm", IsDrp)
End Function

Function AddColoIsRetObj(WiRetSfx As Drs) As Drs
'Fm WiRetSfx : ..RetSfx..
'Ret         : ..IsRetObj @@
Dim IxRetAs%: IxRetAs = IxzAy(WiRetSfx.Fny, "RetSfx")
Dim Dr, Dy(): For Each Dr In Itr(WiRetSfx.Dy)
    Dim RetSfx$: RetSfx = Dr(IxRetAs)
    Dim R As Boolean: R = IsRetObj(RetSfx)
    PushI Dr, R
    PushI Dy, Dr
Next
AddColoIsRetObj = AddColzFFDy(WiRetSfx, "IsRetObj", Dy)
End Function


'
