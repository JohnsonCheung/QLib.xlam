Attribute VB_Name = "MIde_Ty_Mth_ShtTy_Cv"
Option Explicit

Function MthKdzMthTy$(MthTy$)
Select Case MthTy
Case "Function", "Sub": MthKdzMthTy = MthTy
Case "Property Get", "Property Let", "Property Set": MthKdzMthTy = "Property"
End Select
End Function

Function IsMthTy(A$) As Boolean
IsMthTy = HasEle(MthTyAy, A)
End Function

Function IsMthMdy(A$) As Boolean
IsMthMdy = HasEle(MthMdyAy, A)
End Function

Function ShtMthMdy$(A)
Dim O$
Select Case A
Case "Public": O = "Pub"
Case "Private": O = "Prv"
Case "Friend": O = "Frd"
Case ""
Case Else: Stop
End Select
ShtMthMdy = O
End Function
Function MthTySht$(A)
Dim O$
Select Case A
Case "Get": O = "Property Get"
Case "Set": O = "Property Set"
Case "Let": O = "Property Let"
Case "Fun": O = "Function"
Case "Sub": O = "Sub"
End Select
MthTySht = O
End Function

Function MthMdySht$(A)
Dim O$
Select Case A
Case "Pub": O = "Public"
Case "Prv": O = "Private"
Case "Frd": O = "Friend"
Case ""
Case Else: Stop
End Select
MthMdySht = O
End Function

Function ShtMthTy$(A)
Dim O$
Select Case A
Case "Property Get": O = "Get"
Case "Property Set": O = "Set"
Case "Property Let": O = "Let"
Case "Function":     O = "Fun"
Case "Sub":          O = "Sub"
End Select
ShtMthTy = O
End Function
Function ShtMthKdShtMthTy$(A)
Dim O$
Select Case A
Case "Get": O = "Prp"
Case "Set": O = "Prp"
Case "Let": O = "Prp"
Case "Fun": O = "Fun"
Case "Sub": O = "Sub"
End Select
ShtMthKdShtMthTy = O
End Function

Function ShtMthKd$(MthKd)
Dim O$
Select Case MthKd
Case "Property": O = "Prp"
Case "Function": O = "Fun"
Case "Sub":      O = "Sub"
End Select
ShtMthKd = O
End Function

