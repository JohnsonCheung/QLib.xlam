Attribute VB_Name = "MIde_Mth_Lin_RetTy"
Option Explicit

Private Sub Z_MthRetTy()
'Dim MthLin$
'Dim A$:
'MthLin = "Function MthPm(MthPmStr$) As MthPm"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = "MthPm"
'Ass A.IsAy = False
'Ass A.TyChr = ""
'
'MthLin = "Function MthPm(MthPmStr$) As MthPm()"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = "MthPm"
'Ass A.IsAy = True
'Ass A.TyChr = ""
'
'MthLin = "Function MthPm$(MthPmStr$)"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = ""
'Ass A.IsAy = False
'Ass A.TyChr = "$"
'
'MthLin = "Function MthPm(MthPmStr$)"
'A = MthRetTy(MthLin)
'Ass A.TyAsNm = ""
'Ass A.IsAy = False
'Ass A.TyChr = ""
End Sub

Function RetTy$(Lin)
RetTy = MthRetTy(Lin)
End Function
Function MthRetTy$(Lin)
If IsMthLin(Lin) Then MthRetTy = TakMthRetTy(Lin)
End Function
Function TakMthRetTy$(MthLin)
TakMthRetTy = TakBefOrAll(TakAftBkt(MthLin), "'")
End Function
