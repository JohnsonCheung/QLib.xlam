Attribute VB_Name = "MxXlsDta"
Option Explicit
Option Compare Text
Const CLib$ = "QXls."
Const CMod$ = CLib & "MxXlsDta."

Function DrszFxq(Fx, Q) As Drs
DrszFxq = DrszArs(CnzFx(Fx).Execute(Q))
End Function


Function DoFTyzFxw(Fx, W) As Drs
DoFTyzFxw = DoFTy(CatzFx(Fx), CattnzWsn(W))
End Function

Function DoFxwFTy(Fx, W) As Drs
Dim A As Drs
A = DoFTyzFxw(Fx, W)
DoFxwFTy = InsColzDrsCC(A, "Fx W", Fx, W)
End Function


