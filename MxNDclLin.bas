Attribute VB_Name = "MxNDclLin"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxNDclLin."
Public Const FFoNDclLin$ = "Mdn NDclLin"

Function FoNDclLin() As String()
FoNDclLin = SyzSS(FFoNDclLin)
End Function

Function DoNDclLin(P As VBProject) As Drs
DoNDclLin = Drs(FoNDclLin, DyoNDclLin(P))
End Function

Function DyoNDclLin(P As VBProject) As Variant()
Dim C As VBComponent
For Each C In P.VBComponents
    PushI DyoNDclLin, Array(C.Name, NDclLinzM(C.CodeModule))
Next
End Function

Function NDclLinzM%(M As CodeModule) 'Assume FstMth cannot have TopRmk
Dim I&
    I = FstMthLnozM(M)
    If I <= 0 Then
        NDclLinzM = M.CountOfLines
        Exit Function
    End If
NDclLinzM = TopRmkLno(M, I) - 1
End Function

Function NDclLin%(Src$())
Dim Top&
    Dim Fm&
    Fm = FstMthIx(Src)
    If Fm = -1 Then
        NDclLin = UB(Src) + 1
        Exit Function
    End If
NDclLin = IxoPrvCdLin(Src, Fm) + 1
End Function

