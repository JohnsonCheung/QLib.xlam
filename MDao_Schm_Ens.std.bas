Attribute VB_Name = "MDao_Schm_Ens"
Option Explicit

Sub EnsSchm(A As Database, Schm$())
ThwErMsg ErzSchm(Schm), CSub, "there is error in the Schm"
'AppDbTdAy A, TdAy(Smt, AywRmvT1(Schm, CCF), AywRmvT1(Schm, CCE))
'RunSqy A, SqyCrtPk_Tny(PkTnySmt(Smt))
'RunSqy A, SqyCrtSkSmt(Smt)
'Set TblDesDic(A) = TblDesDicSmdt(AywRmvTT(Schm, CCD, CCT))
'Set TblDesDicDb(A) = TblDesDicDbSmdf(AywRmvTT(Schm, CCD, CCF))
End Sub

