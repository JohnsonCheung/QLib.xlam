Attribute VB_Name = "QDao_Db_Get_Dta"
Option Explicit
Private Const CMod$ = "MDao_Db_Get_Dta."
Private Const Asm$ = "QDao"

Function DrzQ(A As Database, Q) As Variant()
DrzQ = DrzRs(Rs(A, Q))
End Function

Function DryzQ(A As Database, Q) As Variant()
DryzQ = DryzRs(Rs(A, Q))
End Function


