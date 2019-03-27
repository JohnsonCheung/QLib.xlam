Attribute VB_Name = "MDao_Db_Get_Dta"
Option Explicit

Function DrzQ(A As Database, Q) As Variant()
DrzQ = DrzRs(Rs(A, Q))
End Function

Function DryzQ(A As Database, Q) As Variant()
DryzQ = DryzRs(Rs(A, Q))
End Function


