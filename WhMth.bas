VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WhMth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim X_Nm As WhNm, X_ShtMdy$(), X_ShtTy$(), _
X_ShtRetTy$(), X_FstArgNm$, X_FstArgShtTy$, X_ArgNy$(), X_ArgSfx$(), X_WiPmOpt As BoolRslt
Dim X_IsEmp As Boolean
Property Get WiPmOpt() As BoolRslt
WiPmOpt = X_WiPmOpt
End Property

Function Init(ShtMdy$(), ShtTy$(), Nm As WhNm) As WhMth
Set X_Nm = Nm
X_ShtMdy = ShtMdy
X_ShtTy = ShtTy
Set X_Nm = Nm
If IsNothing(Nm) Then Thw CSub, "Nm cannot be nothing"
If Nm.IsEmp And (Si(ShtMdy) = 0) And (Si(ShtTy) = 0) Then X_IsEmp = True
Set Init = Me
End Function

Property Get WhNm() As WhNm
Set WhNm = X_Nm
End Property

Property Get IsEmp() As Boolean
IsEmp = X_IsEmp
End Property
Property Get ShtTyAy() As String()
ShtTyAy = X_ShtTy
End Property

Property Get ShtMthMdyAy() As String()
ShtMthMdyAy = X_ShtMdy
End Property

Property Get ToStr$()
If IsEmp Then ToStr = "WhMth(#Emp)": Exit Function
Dim O$()
PushIAy O, AyAddPfx(X_ShtMdy, "-")
PushIAy O, AyAddPfx(X_ShtTy, "-")
ToStr = JnSpc(AyeEmpEle(O))
End Property
