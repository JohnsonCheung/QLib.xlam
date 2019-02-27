VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MthInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public MdNm$, FmLno&, ToLno&, LinCnt%, Lines$, MthLin$, MthNm$, ShtMdy$, ShtKd$, TyChr$, RetTy$, LinRmk$, TopRmk$
Private X_ArgAy$()
Property Get ArgAy() As String()
ArgAy = X_ArgAy
End Property
Property Let ArgAy(V$())
X_ArgAy = V
End Property

