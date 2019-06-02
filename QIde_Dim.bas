Attribute VB_Name = "QIde_Dim"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Dim."
Private Const Asm$ = "QIde"

Function IsDimItmzAs(DimItm$) As Boolean
Dim A$(): A = SyzSS(DimItm)
If Si(A) <> 3 Then Exit Function
If A(1) <> "As" Then Thw CSub, "2nd term is not a [As]", "DimItm", DimItm
If IsNm(A(0)) Then Thw CSub, "1st term is not a name", "DimItm", DimItm
IsDimItmzAs = True
End Function

Function DimNmzSht$(DimShtItm$)
DimNmzSht = RmvChrzSfx(RmvSfxzBkt(DimShtItm), TyChrLis)
End Function

Function DimNmzAs$(DimAsItm$)
DimNmzAs = RmvSfxzBkt(Bef(DimAsItm, " As"))
End Function
Function DimTy$(DimItm$)
Select Case True
Case IsDimItmzSht(DimItm): DimTy = RmvNm(DimItm)
Case IsDimItmzAs(DimItm):  DimTy = Bef(DimItm, " As")
Case Else: Thw CSub, "Not a DimItm", "DimItm", DimItm
End Select
End Function

Function DimNm$(DimItm$)
Select Case True
Case IsDimItmzSht(DimItm): DimNm = DimNmzSht(DimItm)
Case IsDimItmzAs(DimItm):  DimNm = DimNmzAs(DimItm)
Case Else: Thw CSub, "Not a DimItm", "DimItm", DimItm
End Select
End Function

Function IsDimItmzSht(DimItm$) As Boolean
If HasSpc(DimItm) Then Exit Function
IsDimItmzSht = IsNm(RmvTyChr(RmvSfxzBkt(DimItm)))
End Function

Function DimItmAy(Lin) As String()
Dim L$: L = Lin
If Not ShfPfx(L, "Dim ") Then Exit Function
DimItmAy = SplitCommaSpc(L)
End Function

Function DimNyzDimItmAy(DimItmAy$()) As String()
Dim DimItm$, I
For Each I In Itr(DimItmAy)
    DimItm = I
    PushI DimNyzDimItmAy, DimNm(DimItm)
Next
End Function
