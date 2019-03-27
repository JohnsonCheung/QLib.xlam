VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LidFxc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const CMod$ = "LidFxc."
Public ColNm$, ShtTyLis$, ExtNm$
Friend Function Init(ColNm$, ShtTyLis$, ExtNm$) As LidFxc
Const CSub$ = CMod & "Init"
With Me
    .ColNm = ColNm
    .ShtTyLis = ShtTyLis
    .ExtNm = ExtNm
End With
Dim A$(): A = ErzShtTyLis(ShtTyLis)
If Si(A) > 0 Then
    Thw CSub, "Given ShtTyLis has invalid ShtTy", "Invalid-ShtTy Given-ShtTyLis ColNm ExtNm", JnSpc(A), ShtTyLis, ColNm, ExtNm
End If
Set Init = Me
End Function
