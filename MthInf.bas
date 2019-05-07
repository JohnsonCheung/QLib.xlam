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
Private Const CMod$ = "MthInf."
Public MdNm$, FmLno&, ToLno&, LinCnt%, Lines$, MthLin$, MthNm$, ShtMdy$, ShtKd$, TyChr$, RetTy$, LinRmk$, TopRmk$
Private X_ArgSy$()
Property Get ArgSy() As String()
ArgSy = X_ArgSy
End Property
Property Let ArgSy(V$())
X_ArgSy = V
End Property

