Attribute VB_Name = "QXls_Lo_LofVbl"
Option Compare Text
Option Explicit
Private Const CMod$ = "MXls_Lo_LofVbl."
Private Const Asm$ = "QXls"

Function LofVblzQt$(A As QueryTable)
LofVblzQt = LofVblzFbtStr(FbtStrzQt(A))
End Function

Property Get LofVblzT$(A As Database, T)
LofVblzT = TblPrp(A, T, "LofVbl")
End Property

Property Let LofVblzT(A As Database, T, V$)
TblPrp(A, T, "LofVal") = V
End Property

Function LofVblzLo$(A As ListObject)
LofVblzLo = LofVblzQt(LoQt(A))
End Function

Property Get LofVblzFbt$(FB, T)
LofVblzFbt = LofVblzT(Db(FB), T)
End Property

Property Let LofVblzFbt(FB, T, LofVblzVbl$)
LofVblzT(Db(FB), T) = LofVblzVbl
End Property

Function LofVblzFbtStr$(FbtStr$)
Dim FB$, T$
AsgFbtStr FbtStr, FB, T
LofVblzFbtStr = LofVblzFbt(FB, T)
End Function

