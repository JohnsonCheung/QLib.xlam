Attribute VB_Name = "MIde_Ty_Component"
Option Explicit
Const CMod$ = "MIde_Ty_Component."

Function ShtCmpTy$(A As vbext_ComponentType)
Dim O$
Select Case A
Case vbext_ComponentType.vbext_ct_Document:    O = "Doc"
Case vbext_ComponentType.vbext_ct_ClassModule: O = "Cls"
Case vbext_ComponentType.vbext_ct_StdModule:   O = "Std"
Case vbext_ComponentType.vbext_ct_MSForm:      O = "Frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner: O = "ActX"
Case Else: Stop
End Select
ShtCmpTy = O
End Function
Function CmpTyzMd(Md As CodeModule) As vbext_ComponentType
CmpTyzMd = Md.Parent.Type
End Function
Function CmpTy(ShtCmpTy) As vbext_ComponentType
Dim O As vbext_ComponentType
Select Case ShtCmpTy
Case "Doc": O = vbext_ComponentType.vbext_ct_Document
Case "Cls": O = vbext_ComponentType.vbext_ct_ClassModule
Case "Std": O = vbext_ComponentType.vbext_ct_StdModule
Case "Frm": O = vbext_ComponentType.vbext_ct_MSForm
Case "ActX": O = vbext_ComponentType.vbext_ct_ActiveXDesigner
Case Else: Stop
End Select
End Function
