VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiMisTyc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Extnm$, ActShtTy$, EptShtTyLis$
Friend Function Init(Extnm$, ActShtTy$, EptShtTyLis$) As LiMisTyc
With Me
    .Extnm = Extnm
    .ActShtTy = ActShtTy
    .EptShtTyLis = EptShtTyLis
End With
Set Init = Me
End Function

Property Get MisMsg$()
MisMsg = FmtQQ("Column[?] has unexpected-data-type[?], it should be [?]", Extnm, DtaTyzShtTy(ActShtTy), JnSpc(DtaTyAyzShtTyAy(CmlAy(EptShtTyLis))))
End Property

