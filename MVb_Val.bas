Attribute VB_Name = "MVb_Val"
Option Explicit
Function LineszVal$(V)
LineszVal = JnCrLf(FmtV(V))
End Function

Function StrCellzVal$(V, Optional ShwZer As Boolean, Optional MaxWdt%)
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
        O = EscTab(VblzLines(V))
    Else
        O = Left(V, MaxWdt)
        O = Left(EscTab(VblzLines(O)), MaxWdt)
    End If
Case IsStr(V): O = EscTab(V)
Case IsPrim(V): O = V
Case IsSy(V): If Si(V) > 0 Then O = FmtQQ("#Sy(?):", UB(V)) & EscTab(V(0))
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
StrCellzVal = O
End Function
Function FmtLines(Lines$) As String()
FmtLines = FmtSy(SplitCrLf(Lines))
End Function
Function FmtSy(Sy$()) As String()
FmtSy = AyAddIxPfx(Sy, 0)
End Function
Function FmtPrim$(Prim)
FmtPrim = Prim & " (" & TypeName(Prim) & ")"
End Function

Function FmtV(V) As String()
Select Case True
Case IsDic(V): FmtV = FmtDic(CvDic(V))
Case IsAset(V): FmtV = CvAset(V).Sy
Case IsLines(V): FmtV = FmtLines(CStr(V))
Case IsPrim(V): FmtV = Sy(FmtPrim(V))
Case IsSy(V): FmtV = FmtSy(CvSy(V))
Case IsNothing(V): FmtV = Sy("#Nothing")
Case IsEmpty(V): FmtV = Sy("#Empty")
Case IsMissing(V): FmtV = Sy("#Missing")
Case IsObject(V): FmtV = Sy("#Obj(" & TypeName(V) & ")")
Case IsArray(V)
    Dim I, O$()
    If Si(V) = 0 Then Exit Function
    For Each I In V
        PushI O, StrCellzVal(I)
    Next
    FmtV = AyAddIxPfx(O)
Case Else
End Select
End Function

