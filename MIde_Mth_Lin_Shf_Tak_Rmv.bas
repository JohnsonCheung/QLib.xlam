Attribute VB_Name = "MIde_Mth_Lin_Shf_Tak_Rmv"
Option Explicit

Function ShfItmNy(A$, ItmNy0) As Variant()
ShfItmNy = AyShfItmNy(TermAy(A), ItmNy0)
End Function
Function ShfMthTy$(OLin)
Dim O$: O = TakMthTy(OLin)
If O = "" Then Exit Function
ShfMthTy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Sub ShfMthTyAsg(A, OMthTy, ORst$)
AsgAp ShfMthTy(A), OMthTy, ORst
End Sub

Function ShfAs(A) As Variant()
Dim L$
L = LTrim(A)
If Left(L, 3) = "As " Then ShfAs = Array(True, LTrim(Mid(L, 4))): Exit Function
ShfAs = Array(False, A)
End Function
Function ShfShtMthMdy$(OLin)
ShfShtMthMdy = ShtMthMdy(ShfMthMdy(OLin))
End Function
Function ShfShtMthTy$(OLin)
ShfShtMthTy = ShtMthTy(ShfMthTy(OLin))
End Function
Function ShfShtMthKd$(OLin)
ShfShtMthKd = ShtMthKdShtMthTy(ShtMthTy(ShfMthTy(OLin)))
End Function

Function ShfMthMdy$(OLin)
Dim O$
O = TakMthMdy(OLin):
ShfMthMdy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Function ShfMthNm3(OLin) As MthNm3
Set ShfMthNm3 = New MthNm3
Dim M$: M = ShfShtMthMdy(OLin)
Dim T$: T = ShfShtMthTy(OLin):: If T = "" Then Exit Function
ShfMthNm3.Init M, T, ShfNm(OLin)
End Function

Function ShfKd$(OLin)
Dim T$: T = TakMthKd(OLin)
If T = "" Then Exit Function
ShfKd = T
OLin = LTrim(RmvPfx(OLin, T))
End Function

Function ShfMthSfx$(OLin)
ShfMthSfx = ShfChr(OLin, "#!@#$%^&")
End Function

Function ShfNm$(OLin)
Dim O$: O = TakNm(OLin): If O = "" Then Exit Function
ShfNm = O
OLin = RmvPfx(OLin, O)
End Function

Function ShfRmk(A) As String()
Dim L$
L = LTrim(A)
If FstChr(L) = "'" Then
    ShfRmk = Sy(Mid(L, 2), "")
Else
    ShfRmk = Sy("", A)
End If
End Function

Function TakMthMdy$(A)
TakMthMdy = TermLinAy(A, MthMdyAy)
End Function

Function TakMthKd$(A)
TakMthKd = TermLinAy(A, MthKdAy)
End Function

Function TakMthTy$(A)
TakMthTy = TermLinAy(A, MthTyAy)
End Function

Function RmvMdy$(A)
RmvMdy = LTrim(RmvPfxAySpc(A, MthMdyAy))
End Function

Function RmvMthTy$(A)
RmvMthTy = RmvPfxAySpc(A, MthTyAy)
End Function

