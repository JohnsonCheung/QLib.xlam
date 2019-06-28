Attribute VB_Name = "QXls_B_LofVbl"
Option Compare Text
Option Explicit
Private Const CMod$ = "MXls_Lo_LofVbl."
Private Const Asm$ = "QXls"

Function LofVblzQt$(A As QueryTable)
LofVblzQt = LofVblzFbtStr(FbtStrzQt(A))
End Function

Property Get LofVblzT$(D As Database, T)
LofVblzT = TblPrp(D, T, "LofVbl")
End Property

Property Let LofVblzT(D As Database, T, V$)
TblPrp(D, T, "LofVbl") = V
End Property

Function LofVblzLo$(A As ListObject)
LofVblzLo = LofVblzQt(LoQt(A))
End Function

Property Get LofVblzFbt$(Fb, T)
LofVblzFbt = LofVblzT(Db(Fb), T)
End Property

Property Let LofVblzFbt(Fb, T, LofVblzVbl$)
LofVblzT(Db(Fb), T) = LofVblzVbl
End Property

Function LofVblzFbtStr$(FbtStr$)
Dim Fb$, T$
AsgFbtStr FbtStr, Fb, T
LofVblzFbtStr = LofVblzFbt(Fb, T)
End Function

