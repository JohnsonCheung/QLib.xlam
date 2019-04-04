Attribute VB_Name = "MIde_Dim"
Option Explicit

Function IsDimItmzAs(DimItm) As Boolean
Dim A$(): A = SySsl(DimItm)
If Si(A) <> 3 Then Exit Function
If A(1) <> "As" Then Thw CSub, "2nd term is not a [As]", "DimItm", DimItm
If IsNm(A(0)) Then Thw CSub, "1st term is not a name", "DimItm", DimItm
IsDimItmzAs = True
End Function

Function DimNmzSht$(DimShtItm)
DimNmzSht = RmvChrzSfx(RmvSfxzBkt(DimShtItm), TyChrLis)
End Function

Function DimNmzAs$(DimAsItm)
DimNmzAs = RmvSfxzBkt(StrBef(DimAsItm, " As"))
End Function
Function DimTy$(DimItm)
Select Case True
Case IsDimItmzSht(DimItm): DimTy = RmvNm(DimItm)
Case IsDimItmzAs(DimItm):  DimTy = StrBef(DimItm, " As")
Case Else: Thw CSub, "Not a DimItm", "DimItm", DimItm
End Select
End Function

Function DimNm$(DimItm)
Select Case True
Case IsDimItmzSht(DimItm): DimNm = DimNmzSht(DimItm)
Case IsDimItmzAs(DimItm):  DimNm = DimNmzAs(DimItm)
Case Else: Thw CSub, "Not a DimItm", "DimItm", DimItm
End Select
End Function

Function IsDimItmzSht(DimItm) As Boolean
If HasSpc(DimItm) Then Exit Function
IsDimItmzSht = IsNm(RmvTyChr(RmvSfxzBkt(DimItm)))
End Function

Function DimItmAy(Lin) As String()
Dim L$: L = Lin
If Not ShfPfx(L, "Dim ") Then Exit Function
DimItmAy = SplitCommaSpc(L)
End Function

Function DimNy(Lin) As String()
DimNy = DimNyzDimItmAy(DimItmAy(Lin))
End Function

Function DimNyzDimItmAy(DimItmAy$()) As String()
Dim DimItm
For Each DimItm In Itr(DimItmAy)
    PushI DimNyzDimItmAy, DimNm(DimItm)
Next
End Function

Function DimNyzSrc(Src$()) As String()
Dim L
For Each L In Itr(Src)
    PushIAy DimNyzSrc, DimNy(L)
Next
End Function
