Attribute VB_Name = "QVb_Val"
Option Explicit
Private Const CMod$ = "MVb_Val."
Private Const Asm$ = "QVb"
Function LineszV$(V)
LineszV = JnCrLf(FmtV(V))
End Function

Function StrCellzV$(V, Optional ShwZer As Boolean, Optional MaxWdt%)
Dim T$, S$, W%, I, Sep$, O$
Select Case True
Case IsDic(V): O = "#Dic"
Case IsNumeric(V):
    Select Case True
    Case V = 0: If ShwZer Then O = 0
    Case Else: O = V
    End Select
Case IsLines(V):
    If MaxWdt <= 0 Then
        O = SlashTab(VblzLines(CStr(V)))
    Else
        O = Left(V, MaxWdt)
        O = Left(SlashTab(VblzLines(O)), MaxWdt)
    End If
Case IsStr(V): O = SlashTab(CStr(V))
Case IsPrim(V): O = V
Case IsSy(V): If Si(V) > 0 Then O = FmtQQ("#Sy(?):", UB(V)) & SlashTab(CStr(V(0)))
Case IsNothing(V): O = "#Nothing"
Case IsEmpty(V): O = "#Empty"
Case IsMissing(V): O = "#Missing"
Case IsObject(V): O = "#Obj(" & TypeName(V) & ")"
Case IsNull(V): O = "#Null"
Case IsArray(V)
    If Si(V) = 0 Then
        O = "#Emp-" & TypeName(V)
    Else
        O = "#" & TypeName(V) & "(" & UB(V) & ")"
    End If
Case Else
End Select
StrCellzV = O
End Function
Function AddIxPfxzLines(Lines) As String()
AddIxPfxzLines = AddIxPfx(SplitCrLf(Lines))
End Function
Function FmtPrim$(Prim)
FmtPrim = Prim & " (" & TypeName(Prim) & ")"
End Function

Function FmtV(V) As String()
Select Case True
Case IsDic(V): FmtV = FmtDic(CvDic(V))
Case IsAset(V): FmtV = CvAset(V).Sy
Case IsLines(V): FmtV = AddIxPfxzLines(V)
Case IsPrim(V): FmtV = Sy(FmtPrim(V))
Case IsSy(V): FmtV = AddIxPfx(CvSy(V))
Case IsNothing(V): FmtV = Sy("#Nothing")
Case IsEmpty(V): FmtV = Sy("#Empty")
Case IsMissing(V): FmtV = Sy("#Missing")
Case IsObject(V): FmtV = Sy("#Obj(" & TypeName(V) & ")")
Case IsArray(V)
    Dim I, O$()
    If Si(V) = 0 Then Exit Function
    For Each I In V
        PushI O, StrCellzV(I)
    Next
    FmtV = AddIxPfx(O)
Case Else
End Select
End Function

