Attribute VB_Name = "QIde_Ty_Mth_ShtTy_Cv"
Option Explicit
Private Const CMod$ = "MIde_Ty_Mth_ShtTy_Cv."
Private Const Asm$ = "QIde"

Function MthKdByMthTy$(MthTy)
Select Case MthTy
Case "Function", "Sub": MthKdByMthTy = MthTy
Case "Property Get", "Property Let", "Property Set": MthKdByMthTy = "Property"
End Select
End Function

Function IsMthTy(Str$) As Boolean
IsMthTy = HasEle(MthTyAy, Str)
End Function

Function IsMthMdy(A$) As Boolean
IsMthMdy = HasEle(MthMdyAy, A)
End Function

Function MthMdyBySht$(ShtMthMdy)
Dim O$
Select Case ShtMthMdy
Case "Pub": O = "Public"
Case "Prv": O = "Private"
Case "Frd": O = "Friend"
Case ""
Case Else: Stop
End Select
MthMdyBySht = O
End Function

Function ShtMthMdy$(MthMdy)
Dim O$
Select Case MthMdy
Case "Public", "": O = "Pub"
Case "Private": O = "Prv"
Case "Friend": O = "Frd"
Case Else: O = "???"
End Select
ShtMthMdy = O
End Function

Function MthTyBySht$(ShtMthTy)
Dim O$
Select Case ShtMthTy
Case "Get": O = "Property Get"
Case "Set": O = "Property Set"
Case "Let": O = "Property Let"
Case "Fun": O = "Function"
Case "Sub": O = "Sub"
Case Else: O = "???"
End Select
MthTyBySht = O
End Function

Function ShtMthTy$(MthTy)
Dim O$
Select Case MthTy
Case "Property Get": O = "Get"
Case "Property Set": O = "Set"
Case "Property Let": O = "Let"
Case "Function":     O = "Fun"
Case "Sub":          O = "Sub"
Case Else: O = "???"
End Select
ShtMthTy = O
End Function
Private Sub Z_ShtMthTyzLin()
GoSub ZZ
Exit Sub
ZZ:
    Dim O$(), I, Lin$
    For Each I In MthLinSyInVbe
        Lin = I
        PushI O, ShtMthTyzLin(Lin)
    Next
    Brw O
    Return
End Sub
Function ShtMthTyzLin$(Lin$)
ShtMthTyzLin = ShtMthTy(TakMthTy(RmvMthMdy(Lin)))
End Function

Function ShtMthKdzShtMthTy$(ShtMthTy$)
Dim O$
Select Case ShtMthTy
Case "Get": O = "Prp"
Case "Set": O = "Prp"
Case "Let": O = "Prp"
Case "Fun": O = "Fun"
Case "Sub": O = "Sub"
Case Else: O = "???"
End Select
ShtMthKdzShtMthTy = O
End Function

Function ShtMthKd$(MthKd)
Dim O$
Select Case MthKd
Case "Property": O = "Prp"
Case "Function": O = "Fun"
Case "Sub":      O = "Sub"
Case Else: O = "???"
End Select
ShtMthKd = O
End Function

