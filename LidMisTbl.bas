VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LidMisTbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const CMod$ = "LidMisTbl."
Public Ffn$, T$, Wsn$, FilNm$
Friend Function Init(Ffn, FilNm, T, Optional Wsn$) As LidMisTbl
Const CSub$ = CMod & "Init"
Dim M$
With Me
    .FilNm = FilNm
    .Ffn = Ffn
    .T = T
    Select Case True
    Case IsFx(Ffn): If Wsn = "" Then M = "Wsn must be given when Fx": GoTo X
    Case IsFb(Ffn): If Wsn <> "" Then M = "Wsn must not be given when Fb": GoTo X
    Case Else: M = "Ffn must be Fx or Fb": GoTo X
    End Select
    .Wsn = Wsn
End With
Set Init = Me
Exit Function
X: ThwNav CSub, M, Av("Ffn FilNm T Wsn", Ffn, FilNm, T, Wsn)
End Function

Property Get MisMsg$()
MisMsg = FmtQQ("In ?-?, ?[?] is missing, in folder[?], file-name[?]", _
    FilNm, FfnKd(Ffn), TblKd(Ffn), T, Pth(Ffn), Fn(Ffn))
End Property



