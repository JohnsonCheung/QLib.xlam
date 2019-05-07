Attribute VB_Name = "QDao_Schm_Ens"
Option Explicit
Private Const CMod$ = "MDao_Schm_Ens."
Private Const Asm$ = "QDao"

Sub EnsSchm(A As Database, Schm$())
Stop
ThwIfErMsg ErzSchm(Schm), CSub, "there is error in the Schm"
'AppDbTdAy A, TdAy(Smt, SywRmvT1(Schm, CCF), SywRmvT1(Schm, CCE))
'RunSqy A, SqyCrtPk_Tny(PkTnySmt(Smt))
'RunSqy A, SqyCrtSkSmt(Smt)
'Set TblDesDic(A) = TblDesDicSmdt(SywRmvTT(Schm, CCD, CCT))
'Set TblDesDicDb(A) = TblDesDicDbSmdf(SywRmvTT(Schm, CCD, CCF))
End Sub

