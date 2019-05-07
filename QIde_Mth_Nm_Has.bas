Attribute VB_Name = "QIde_Mth_Nm_Has"
Option Explicit
Private Const CMod$ = "MIde_Mth_Nm_Has."
Private Const Asm$ = "QIde"
Function HasMthSrc(Src$(), MthNm) As Boolean
HasMthSrc = MthIxzFst(Src, MthNm, 0) >= 0
End Function

Function HasMthMd(A As CodeModule, MthNm) As Boolean
HasMthMd = HasMthSrc(Src(A), MthNm)
End Function
