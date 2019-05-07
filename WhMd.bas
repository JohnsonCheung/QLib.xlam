VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WhMd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const CMod$ = "WhMd."
Public Nm As WhNm
Dim X_CmpTy() As vbext_ComponentType
Function Init(CmpTy() As vbext_ComponentType, Nm As WhNm) As WhMd
X_CmpTy = CmpTy
Set Me.Nm = Nm
Set Init = Me
End Function
Property Get CmpTy() As vbext_ComponentType()
CmpTy = X_CmpTy
End Property
