Attribute VB_Name = "QDao_Db_Get_Dta"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Db_Get_Dta."
Private Const Asm$ = "QDao"

Function DrzQ(A As Database, Q) As Variant()
DrzQ = DrzRs(Rs(A, Q))
End Function

Function DyoQ(A As Database, Q) As Variant()
DyoQ = DyoRs(Rs(A, Q))
End Function


