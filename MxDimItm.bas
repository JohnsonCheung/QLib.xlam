Attribute VB_Name = "MxDimItm"
Option Explicit
Option Compare Text
Const CLib$ = "QIde."
Const CMod$ = CLib & "MxDimItm."
Sub Z_DimItmAy()
Brw DimItmAy(SrczP(CPj))
End Sub

Function DimItmAy(Src$()) As String()
':DimItm: :S #Dim-Itm#
DimItmAy = DimItmAyzDimLinAy(DimLinAy(Src))
End Function
Sub Z_DimLinAy()
Brw DimLinAy(SrczP(CPj))
End Sub

Function DimLinAy(Src$()) As String()
'DimLin: :Lin #Dim-Lin# ! the Fst4Chr must be [Dim ]
Dim L: For Each L In Itr(Src)
    L = LTrim(L)
    If FstChr(L) <> "'" Then
        Dim P%: P = InStr(L, "Dim ")
        If P > 0 Then
            If InStr(L, """Dim") = 0 Then
                PushI DimLinAy, BefOrAll(BefOrAll(Mid(L, P), ":"), "'")
            End If
        End If
    End If
Next
End Function

Function DimItmAyzDimLinAy(DimLinAy$()) As String()
Dim DimLin: For Each DimLin In Itr(DimLinAy)
    If Left(DimLin, 4) <> "Dim " Then Stop
    PushIAy DimItmAyzDimLinAy, AmTrim(SplitComma(Mid(DimLin, 5)))
Next
End Function
