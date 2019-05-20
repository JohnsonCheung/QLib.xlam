Attribute VB_Name = "QIde_Mod_ChgToCmp"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Mod_ChgToCmp."
Private Const Asm$ = "QIde"
Sub ChgToCls(FmModn$)
If Not HasCmp(FmModn) Then InfLin CSub, "Mod not exist", "Mod", FmModn: Exit Sub
If Not IsMod(Md(FmModn)) Then InfLin CSub, "It not Mod", "Mod", FmModn: Exit Sub
Dim T$: T = Left(FmModn & "_" & Format(Now, "HHMMDD"), 31)
Md(FmModn).Name = T
AddCls FmModn
Md(FmModn).AddFromString SrcLines(Md(T))
RmvCmpzN T
End Sub
