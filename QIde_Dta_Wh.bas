Attribute VB_Name = "QIde_Dta_Wh"
Type WhMth
    WhNm As WhNm
    ShtMdy() As String
    ShtTy() As String
    ShtRetTy() As String
    FstArgNm As String
    FstArgShtTy As String
    ArgNy() As String
    ArgSfx() As String
    WiPmOpt As BoolOpt
End Type
Type WhMd
    CmpTy() As vbext_ComponentType
    WhNm As WhNm
End Type

Type WhMdMth
    WhMd As WhMd
    WhMth As WhMth
End Type
Type WhPjMth
    WhPjNm As WhNm
    WhMdMth As WhMdMth
End Type
Function WhMth(ShtMdy$(), ShtTy$(), Nm As WhNm) As WhMth
With WhMth
    .WhNm = Nm
.ShtMdy = ShtMdy
.ShtTy = ShtTy
'If IsNothing(Nm) Then Thw CSub, "Nm cannot be nothing"
'If Nm.IsEmp And (Si(ShtMdy) = 0) And (Si(ShtTy) = 0) Then X_IsEmp = True
End With
End Function

Function WhMthStr$(A As WhMth)
'If IsEmpWhMth(A) Then WhMthNmStr = "WhMth(#Emp)": Exit Function
Dim O$()
'PushIAy O, AddPfxzAy(X_ShtMdy, "-")
'PushIAy O, AddPfxzAy(X_ShtTy, "-")
'ToStr = JnSpc(AyeEmpEle(O))
End Function

