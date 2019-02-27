VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LidMisTyc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ExtNm$, ActShtTy$, EptShtTyLis$
Friend Function Init(ExtNm, ActShtTy$, EptShtTyLis$) As LidMisTyc
With Me
    .ExtNm = ExtNm
    .ActShtTy = ActShtTy
    .EptShtTyLis = EptShtTyLis
End With
Set Init = Me
End Function

Property Get MisMsg$()
MisMsg = FmtQQ("Column[?] has unexpected-data-type[?], it should be [?]", ExtNm, DtaTyzShtTy(ActShtTy), JnOr(DtaTyAyzShtTyAy(CmlAy(EptShtTyLis))))
End Property


