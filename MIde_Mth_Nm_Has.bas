Attribute VB_Name = "MIde_Mth_Nm_Has"
Option Explicit
Function HasMthSrc(Src$(), MthNm) As Boolean
HasMthSrc = MthIxzFst(Src, MthNm, 0) >= 0
End Function

Function HasMthMd(A As CodeModule, MthNm) As Boolean
HasMthMd = HasMthSrc(Src(A), MthNm)
End Function
