Attribute VB_Name = "MDao_Db_Get_Dta"
Option Explicit

Function DrzDbq(A As Database, Q) As Variant()
DrzDbq = DrzRs(Rsz(A, Q))
End Function

Function DrzQ(Q) As Variant()
DrzQ = DrzDbq(CDb, Q)
End Function

Function Drsz(Q) As Drs
Set Drsz = DrszDbq(CDb, Q)
End Function

