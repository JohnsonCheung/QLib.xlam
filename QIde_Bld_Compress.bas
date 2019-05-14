Attribute VB_Name = "QIde_Bld_Compress"
Private Sub Z_CompressFxa()
CompressFxa Pjf(CPj)
End Sub

Sub CompressFxa(Fxa$)
'PjExp PjzPjf(Xls.Vbe, Fxa)
Dim Srcp$: Srcp = SrcpzPjf(Fxa)
'CrtDistFxa Srcp
RplFfn Fxa, Srcp
End Sub

