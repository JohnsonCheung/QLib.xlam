Attribute VB_Name = "QVb_Str_Val"
Option Compare Text
Option Explicit
Private Const CMod$ = "MVb_Val."
Private Const Asm$ = "QVb"

Function AddIxPfxzLines(Lines, Optional B As EmIxCol = EiBeg0) As String()
AddIxPfxzLines = AddIxPfx(SplitCrLf(Lines), B)
End Function

Function FmtPrim$(Prim)
FmtPrim = Prim & " (" & TypeName(Prim) & ")"
End Function

Function FmtV(V, Optional IsAddIx As Boolean) As String()
Select Case True
Case IsDic(V): FmtV = FmtDic(CvDic(V))
Case IsAset(V): FmtV = CvAset(V).Sy
Case IsLines(V): FmtV = AddIxPfxzLines(V)
Case IsPrim(V): FmtV = Sy(FmtPrim(V))
Case IsSy(V)
    If IsAddIx Then
        FmtV = AddIxPfx(CvSy(V))
    Else
        FmtV = V
    End If
Case IsNothing(V): FmtV = Sy("#Nothing")
Case IsEmpty(V): FmtV = Sy("#Empty")
Case IsMissing(V): FmtV = Sy("#Missing")
Case IsObject(V): FmtV = Sy("#Obj(" & TypeName(V) & ")")
Case IsArray(V)
    Dim I, O$()
    If Si(V) = 0 Then Exit Function
    For Each I In V
        PushI O, Cell(I)
    Next
    If IsAddIx Then
        FmtV = AddIxPfx(O)
    Else
        FmtV = O
    End If
Case Else
End Select
End Function


'
