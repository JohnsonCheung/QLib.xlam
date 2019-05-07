Attribute VB_Name = "QIde_Mth_Lin_RetTy"
Option Explicit
Private Const CMod$ = "MIde_Mth_Lin_RetTy."
Private Const Asm$ = "QIde"

Private Sub Z_MthRetTy()
'Dim MthLin$
'Dim A$:
'MthLin = "Function MthPm(MthPm$) As MthPm"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = "MthPm"
'Ass A.IsAy = False
'Ass A.TyChr = ""
'
'MthLin = "Function MthPm(MthPm$) As MthPm()"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = "MthPm"
'Ass A.IsAy = True
'Ass A.TyChr = ""
'
'MthLin = "Function MthPm$(MthPm$)"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = ""
'Ass A.IsAy = False
'Ass A.TyChr = "$"
'
'MthLin = "Function MthPm(MthPm$)"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = ""
'Ass A.IsAy = False
'Ass A.TyChr = ""
End Sub

Function MthRetTy$(Lin)
If IsMthLin(Lin) Then MthRetTy = BefOrAll(AftBkt(Lin), "'")
End Function
