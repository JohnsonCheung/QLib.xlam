Attribute VB_Name = "QIde_Cmd_LisSrc"
Option Compare Text
Option Explicit
Private Const CMod$ = "MIde_Cmd_Lis_Src."
Private Const Asm$ = "QIde"
Sub SLis(Patn$)
SLiszPP CPj, Patn
End Sub
Sub BrwzPP(Pj As VBProject, Patn$)
Brw SLocyzPP(Pj, Patn)
End Sub
Sub SLiszPP(P As VBProject, Patn$) 'SLis = SrcLis
D SLocyzPP(P, Patn$)
End Sub


