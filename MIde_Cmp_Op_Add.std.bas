Attribute VB_Name = "MIde_Cmp_Op_Add"
Option Explicit
Function AddCmpM(Nm) As VBComponent
Set AddCmpM = AddCmp(Nm, vbext_ct_StdModule)
End Function
Function AddCmpC(Nm) As VBComponent
Set AddCmpC = AddCmp(Nm, vbext_ct_ClassModule)
End Function
Function AddCmp(Nm, Ty As vbext_ComponentType) As VBComponent
Set AddCmp = AddCmpzPj(CurPj, Nm, Ty)
End Function
Function AddCmpzPj(A As VBProject, Nm, Ty As vbext_ComponentType) As VBComponent
If HasCmp(Nm) Then InfoLin CSub, FmtQQ("?[?] already exist", ShtCmpTy(Ty), Nm): Exit Function
Dim O As VBComponent
Set O = A.VBComponents.Add(Ty)
O.Name = CStr(Nm) ' no CStr will break
Set AddCmpzPj = O
End Function

