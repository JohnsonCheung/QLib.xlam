Attribute VB_Name = "MIde_Mod_ChgToCmp"
Option Explicit
Sub ChgToCmpz(FmModNm$)
If Not HasCmp(FmModNm) Then InfLin CSub, "Mod not exist", "Mod", FmModNm: Exit Sub
If Not IsMod(Md(FmModNm)) Then InfLin CSub, "It not Mod", "Mod", FmModNm: Exit Sub
Dim T$: T = Left(FmModNm & "_" & Format(Now, "HHMMDD"), 31)
Md(FmModNm).Name = T
AddCmp(FmModNm, vbext_ct_ClassModule).CodeModule.AddFromString SrcLines(Md(T))
End Sub

Sub ChgToCmp()
ChgToCmpz CurMdNm
End Sub
