Attribute VB_Name = "MVb_Rslt"
Option Explicit
Type LyRslt: Er() As String: Ly() As String: End Type
Type StrRslt: Som As Boolean: Str As String: End Type
Type BoolRslt: Som As Boolean: Bool As Boolean: End Type
Type DicRslt: Som As Boolean: Dic As Dictionary: End Type
Type LngRslt: Som As Boolean: Lng As Long: End Type
Function LngRslt(Lng) As LngRslt: LngRslt.Som = True: LngRslt.Lng = Lng: End Function
Function LyRslt(Er$(), Ly$()) As LyRslt: LyRslt.Er = Er: LyRslt.Ly = Ly: End Function
Function StrRslt(Str) As StrRslt: StrRslt.Str = Str: StrRslt.Som = True: End Function
Function BoolRslt(Bool As Boolean) As BoolRslt: BoolRslt.Bool = Bool: BoolRslt.Som = True: End Function
Function DicRslt(Dic As Dictionary) As DicRslt: Set DicRslt.Dic = Dic: DicRslt.Som = True: End Function
Function TrueRslt() As BoolRslt: TrueRslt = BoolRslt(True): End Function
Function FalseRslt() As BoolRslt: FalseRslt = BoolRslt(False): End Function

