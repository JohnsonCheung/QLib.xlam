Attribute VB_Name = "MXls_Lo_LofVbl"
Option Explicit

Function LofVblzQt$(A As QueryTable)
LofVblzQt = LofVblzFbtStr(FbtStrQt(A))
End Function

Property Get LofVblzTbl$(T)
LofVblzTbl = LofVblzDbt(CDb, T)
End Property

Property Let LofVblzTbl(T, LofVblzVbl$)
LofVblzDbt(CDb, T) = LofVblzVbl
End Property

Property Get LofVblzDbt$(A As Database, T)
'LofVblzDbt = Dbt(A, T).Prp("LofVblz")
End Property
Property Let LofVblzDbt(A As Database, T, V$)
'Dbt(A, T).Prp("LofVblz") = V
End Property

Function LofVblzLo$(A As ListObject)
LofVblzLo = LofVblzQt(LoQt(A))
End Function

Property Get LofVblzFbt$(Fb, T)
LofVblzFbt = LofVblzDbt(Db(Fb), T)
End Property

Property Let LofVblzFbt(Fb, T, LofVblzVbl$)
LofVblzDbt(Db(Fb), T) = LofVblzVbl
End Property

Function LofVblzFbtStr$(FbtStr$)
Dim Fb$, T$
AsgFbtStr FbtStr, Fb, T
LofVblzFbtStr = LofVblzFbt(Fb, T)
End Function

