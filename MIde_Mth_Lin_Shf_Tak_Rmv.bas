Attribute VB_Name = "MIde_Mth_Lin_Shf_Tak_Rmv"
Option Explicit

Function ShfMthTy$(OLin$)
Dim O$: O = TakMthTy(OLin)
If O = "" Then Exit Function
ShfMthTy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Function ShfTermAftAs$(OLin$)
If Not ShfTermX(OLin, "As") Then Exit Function
ShfTermAftAs = ShfT1(OLin)
End Function
Function ShfShtMthMdy$(OLin$)
ShfShtMthMdy = ShtMthMdy(ShfMthMdy(OLin))
End Function
Function ShfShtMthTy$(OLin$)
ShfShtMthTy = ShtMthTy(ShfMthTy(OLin))
End Function
Function ShfShtMthKd$(OLin$)
ShfShtMthKd = ShtMthKdzShtMthTy(ShtMthTy(ShfMthTy(OLin)))
End Function

Function ShfMthMdy$(OLin$)
Dim O$
O = MthMdy(OLin):
ShfMthMdy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Function ShfMthNm3(OLin$) As MthNm3
Set ShfMthNm3 = New MthNm3
Dim M$: M = ShfShtMthMdy(OLin)
Dim T$: T = ShfShtMthTy(OLin):: If T = "" Then Exit Function
ShfMthNm3.Init M, T, ShfNm(OLin)
End Function

Function ShfKd$(OLin$)
Dim T$: T = TakMthKd(OLin)
If T = "" Then Exit Function
ShfKd = T
OLin = LTrim(RmvPfx(OLin, T))
End Function

Function ShfMthSfx$(OLin$)
ShfMthSfx = ShfChr(OLin, "#!@#$%^&")
End Function

Function ShfNm$(OLin$)
Dim O$: O = Nm(OLin): If O = "" Then Exit Function
ShfNm = O
OLin = RmvPfx(OLin, O)
End Function

Function ShfRmk$(OLin$)
Dim L$
L = LTrim(OLin)
If FstChr(L) = "'" Then
    ShfRmk = Mid(L, 2)
    OLin = ""
End If
End Function

Function TakMthKd$(S$)
TakMthKd = PfxAySpc(S, MthKdAy)
End Function

Function TakMthTy$(S$)
TakMthTy = PfxAySpc(S, MthTyAy)
End Function

Function RmvMdy$(S$)
RmvMdy = LTrim(RmvPfxAySpc(S$, MthMdyAy))
End Function

Function RmvMthTy$(S$)
RmvMthTy = RmvPfxAySpc(S, MthTyAy)
End Function

