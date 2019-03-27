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
Dim X_Nm As WhNm, X_ShtMdy$(), X_ShtKd$()
Dim X_IsEmp As Boolean

Function Init(ShtMdy$(), ShtKd$(), Nm As WhNm) As WhMth
Set X_Nm = Nm
X_ShtMdy = ShtMdy
X_ShtKd = ShtKd
Set X_Nm = Nm
If IsNothing(Nm) Then Thw CSub, "Nm cannot be nothing"
If Nm.IsEmp And (Si(ShtMdy) = 0) And (Si(ShtKd) = 0) Then X_IsEmp = True
Set Init = Me
End Function

Property Get WhNm() As WhNm
Set WhNm = X_Nm
End Property
Property Get IsEmp() As Boolean
IsEmp = X_IsEmp
End Property
Property Get ShtKdAy() As String()
ShtKdAy = X_ShtKd
End Property

Property Get ShtMdyAy() As String()
ShtMdyAy = X_ShtMdy
End Property

Property Get ToStr$()
If IsEmp Then ToStr = "WhMth(#Emp)": Exit Function
Dim O$()
PushI O, AyAddPfx(X_ShtMdy, "-")
PushI O, AyAddPfx(X_ShtKd, "-")
ToStr = JnSpc(AyeEmpEle(O))
End Property
