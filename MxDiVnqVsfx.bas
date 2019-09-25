Attribute VB_Name = "MxDiVnqVsfx"
Option Explicit
Option Compare Text
Const CMod$ = CLib & "MxDiVnqVsfx."
Function DiVnqVsfx(Mthly$()) As Dictionary
Dim L: For Each L In Itr(Mthly)
    PushDiS12 DiVnqVsfx, S12oVnqVsfx(DimItmAyzS(Mthly))
Next
End Function

Function S12oVnqVsfx(DimItm) As S12
S12oVnqVsfx.S1 = TakNm(DimItm)
S12oVnqVsfx.S2 = Vsfx(AftNm(DimItm))
End Function

Function Vsfx$(AftNm_OfDimItm$)
Dim S$: S = LTrim(AftNm_OfDimItm)
Select Case True
Case S = "": Exit Function
Case Fst2Chr(S) = "()"
    If Len(S) = 2 Then
        Vsfx = S
    Else
        S = LTrim(Mid(S, 3))
        If HasPfx(S, "As ") Then
            Vsfx = ":" & Trim(RmvPfx(S, "As ")) & "()"
        Else
            Thw CSub, "Invalid AftNm_OfDimItm", "When aft :() , it should be :As", "AftNm_OfDimItm", AftNm_OfDimItm
        End If
    End If
Case HasPfx(S, "As ")
    Vsfx = ":" & RmvPfx(RmvPfx(S, "As "), "New ")
Case Else
    Vsfx = S
End Select
End Function

Function S12soVnqVsfxP() As S12s
S12soVnqVsfxP = S12soVnqVsfxzP(CPj)
End Function

Function S12soVnqVsfxzP(P As VBProject) As S12s
S12soVnqVsfxzP = S12soVnqVsfx(DimItmAyzS(SrczP(P)))
End Function

Function S12soVnqVsfx(DimItmAy$()) As S12s
Dim I: For Each I In Itr(DimItmAy)
    PushS12 S12soVnqVsfx, S12oVnqVsfx(I)
Next
End Function
