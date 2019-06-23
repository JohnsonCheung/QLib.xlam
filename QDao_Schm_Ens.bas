Attribute VB_Name = "QDao_Schm_Ens"
Option Compare Text
Option Explicit
Private Const CMod$ = "MDao_Schm_Ens."
Private Const Asm$ = "QDao"

Sub EnsSchm(A As Database, Schm$())
Stop
ThwIf_ErMsg ErzSchm(Schm), CSub, "there is error in the Schm"
'AppDbTdAy A, TdAy(Smt, AwRmvT1(Schm, CCF), AwRmvT1(Schm, CCE))
'RunSqy A, SqyCrtPk_Tny(PkTnySmt(Smt))
'RunSqy A, SqyCrtSkSmt(Smt)
'Set TblDesDic(A) = TblDesDicSmdt(AwRmvTT(Schm, CCD, CCT))
'Set TblDesDicDb(A) = TblDesDicDbSmdf(AwRmvTT(Schm, CCD, CCF))
End Sub

