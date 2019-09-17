Attribute VB_Name = "MxAdoDta"
Option Explicit
Option Compare Text
Const CLib$ = "QDao."
Const CMod$ = CLib & "MxAdoDta."
Function DrsCnq(Cn As ADODB.Connection, Q) As Drs
DrsCnq = DrszArs(ArszCnq(Cn, Q))
End Function

Function DrsFbqAdo(Fb, Q) As Drs
DrsFbqAdo = DrszArs(ArszFbq(Fb, Q))
End Function

Function DrszArs(A As ADODB.Recordset) As Drs
DrszArs = Drs(FnyzArs(A), DyoArs(A))
End Function

